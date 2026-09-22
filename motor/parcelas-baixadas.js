// Relatório das parcelas lançadas no Totvs que foram baixadas.
//
//   node motor/parcelas-baixadas.js <financeiro.xlsx> [pasta-de-saida] //        [--totvs a.xlsx --unidades b.xlsx --mkt c.xlsx] [--so reserva]
//
// --so reserva limita o relatorio as parcelas de Reserva de vaga e grava com
// outro nome, sem tocar no relatorio completo.
//
// Com as tres bases extras o relatorio ganha a coluna Canal: quem pagou pelo
// Marketplace e quem pagou pelo Portal do aluno.
//
// Baixada = VALOR_BAIXADO maior que zero com DATA_BAIXA válida. A coluna
// DATA_BAIXA traz lixo numérico na maioria das linhas, então os dois sinais
// precisam concordar.

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');
const { pasta, colName } = require('./escrever-xlsx.js');

const argv = process.argv.slice(2);
const flags = {}, pos = [];
for (let i = 0; i < argv.length; i++) {
  if (argv[i].startsWith('--')) flags[argv[i].slice(2)] = argv[++i];
  else pos.push(argv[i]);
}
const [pFin, SAIDA = '.'] = pos;
if (!pFin) {
  console.error('uso: node parcelas-baixadas.js <financeiro.xlsx> [pasta-de-saida]');
  process.exit(1);
}

const CODE = {
  BG:'Bangu', CG:'Campo Grande', CX:'Duque de Caxias', MD:'Madureira', NI:'Nova Iguaçu',
  RM:'Rocha Miranda', RT:'Retiro dos Artistas', SJ:'São João de Meriti', TQ:'Taquara',
  TJ:'Tijuca', AM:'Américas'
};
const ORDER = ['BG','CG','AM','CX','RM','TQ','TJ','NI','MD','RT','SJ'];
const DE_FILIAL = {
  'TAQUARA':'TQ','DUQUE DE CAXIAS':'CX','BANGU':'BG','ROCHA MIRANDA':'RM','NOVA IGUACU':'NI',
  'SÃO JOÃO DE MERITI':'SJ','CAMPO GRANDE':'CG','RETIRO DOS ARTISTAS':'RT','TIJUCA':'TJ',
  'MADUREIRA':'MD','RECREIO':'AM'
};
const unidade = f => DE_FILIAL[
  String(f||'').replace(/COL[ÉE]GIO E CURSO MATRIZ EDUCA[ÇC][ÃA]O( -)? /,'').trim()] || '??';

// O export financeiro nao traz quem operou a baixa. USUARIO RESP. MATRICULA e o
// usuario dono do registro de matricula - constante em todas as linhas do RA,
// inclusive nas que nunca foram baixadas. E o que da para rastrear hoje.
const SISTEMA = { p_worknow_int_hb:'Integração Marketplace', IMPORTADOR:'Importação automática',
  mestre:'Usuário mestre', totvs:'Usuário Totvs' };
const MINUSC = new Set(['de','da','do','das','dos','e']);
const pessoa = v => {
  const u = String(v == null ? '' : v).trim();
  if (!u) return '';
  if (SISTEMA[u]) return SISTEMA[u];
  if (!/[._]/.test(u) && u === u.toUpperCase()) return u;
  return u.split(/[._]/).filter(Boolean).map((p, i) =>
    (i && MINUSC.has(p.toLowerCase())) ? p.toLowerCase()
      : p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(' ');
};

const num = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; };
const ehData = v => /^\d{2}\/\d{2}\/\d{4}$/.test(String(v || '').trim());
const paraISO = v => ehData(v) ? String(v).trim().split('/').reverse().join('-') : null;
// Parcelada = a entrada foi dividida em mais de uma cobrança. Só vale para a
// Reserva de vaga, onde o contrato é a própria entrada: no serviço 1ª Cota de
// Mensalidade o plano de 12 parcelas é a anuidade, e a cota continua sendo uma
// parcela só. No Marketplace o parcelamento é do cartão e a escola recebe
// integral, então também não entra aqui.
const ehParcelado = r => /Reserva de vaga/i.test(String(r[34] || '')) &&
  ((Number(String(r[27] || '1').replace(',', '.')) || 1) > 1 ||
   !/^\s*PARCELA UNICA\s*$/i.test(String(r[25] || 'PARCELA UNICA')));


// Canal do pagamento. O pareamento Marketplace <-> Totvs e o do motor: ele
// resolve 1 para 1, o que importa aqui porque ha irmaos dividindo a mesma
// transacao e alunos com uma cobranca estornada e outra paga.
const canalPorRA = new Map();
const COM_CANAL = !!(flags.totvs && flags.unidades && flags.mkt);
if (COM_CANAL) {
  const { cruzar } = require('./cruzar.js');
  const r = cruzar(flags.totvs, flags.unidades, pFin === flags.mkt ? null : flags.mkt, pFin);
  r.itens.forEach(i => { if (!i.ra) return;
    canalPorRA.set(i.ra, i.pagoM ? 'Marketplace'
      : i.pg === 'Estornado' ? 'Portal (estornado no Marketplace)'
      : 'Portal do aluno'); });
}
// Rematricula nao passa pelo Marketplace, que so vende bolsao.
const canal = ra => canalPorRA.get(ra) || 'Portal do aluno';

// Recorte por servico. Hoje so "reserva"; o relatorio cheio continua o padrao.
const SO_RESERVA = /^reserva/i.test(String(flags.so || ''));
const RESERVA = /Reserva de vaga/i;
const TITULO = SO_RESERVA ? 'Reservas de vaga baixadas no Totvs' : 'Parcelas baixadas no Totvs';
const ARQUIVO = SO_RESERVA ? 'Reservas de vaga baixadas no Totvs - 2027.xlsx'
                           : 'Parcelas baixadas no Totvs - 2027.xlsx';

const { abas } = lerPlanilha(pFin);
const L = abas.filter(a => a.linhas).sort((a, b) => b.linhas.length - a.linhas.length)[0].linhas;
const linhas = L.slice(1).filter(r => r && r[3]);

// O export repete a mesma parcela quando o boleto tem dois status de remessa:
// duas linhas com o mesmo REF_FINANCEIRO, valor e data, uma "Em aberto" e
// outra "Baixado". E um so pagamento. Fico com a linha ja baixada.
const semRepetir = rs => {
  const porRef = new Map();
  for (const r of rs) {
    const ref = String(r[35] || '').trim();
    if (!ref) { porRef.set('sem-ref-' + porRef.size, r); continue; }
    const atual = porRef.get(ref);
    if (!atual || (String(r[47]) === 'Baixado' && String(atual[47]) !== 'Baixado'))
      porRef.set(ref, r);
  }
  return [...porRef.values()];
};

const baixadas = semRepetir(linhas.filter(r => num(r[44]) > 0 && ehData(r[45]) &&
    (!SO_RESERVA || RESERVA.test(String(r[34] || '')))))
  .map(r => ({
    code: unidade(r[2]), filial: r[2], ra: r[3], aluno: String(r[4] || '').trim(),
    resp: r[7] || null, cpfResp: r[8] || null,
    tipo: r[19], situacao: r[21], curso: r[15], serie: r[16], descricao: r[17], turno: r[18],
    contrato: r[26], servico: r[34], parcela: r[28], cota: r[29],
    original: num(r[42]), deducoes: num(r[43]), baixado: num(r[44]),
    dtBaixa: paraISO(r[45]), venc: paraISO(r[41]), boleto: r[47],
    login: String(r[13] || '').trim(), usuario: pessoa(r[13]),
    plano: r[25] || null, canal: canal(r[3]), reserva: /Reserva de vaga/i.test(String(r[34] || '')),
    qtdPlano: r[27] || null, parcelado: ehParcelado(r)
  }))
  .sort((a, b) => ORDER.indexOf(a.code) - ORDER.indexOf(b.code) ||
                  a.aluno.localeCompare(b.aluno, 'pt') ||
                  String(a.dtBaixa).localeCompare(String(b.dtBaixa)));

/* ---- uma linha por aluno ---- */
const porAluno = new Map();
for (const p of baixadas) {
  if (!porAluno.has(p.ra)) porAluno.set(p.ra, {
    code: p.code, ra: p.ra, aluno: p.aluno, resp: p.resp, tipo: p.tipo,
    situacao: p.situacao, curso: p.curso, serie: p.serie, descricao: p.descricao,
    canal: p.canal, parcelado: false, planos: new Set(),
    n: 0, total: 0, primeira: null, ultima: null,
    servicos: new Set(), usuarios: new Set()
  });
  const a = porAluno.get(p.ra);
  a.n++; a.total += p.baixado; if (p.usuario) a.usuarios.add(p.usuario);
  if (p.parcelado) a.parcelado = true;
  if (p.plano) a.planos.add(p.plano); a.servicos.add(String(p.servico || '').replace(/^\w+ - /, ''));
  if (!a.primeira || p.dtBaixa < a.primeira) a.primeira = p.dtBaixa;
  if (!a.ultima || p.dtBaixa > a.ultima) a.ultima = p.dtBaixa;
}
const alunos = [...porAluno.values()].sort((a, b) =>
  ORDER.indexOf(a.code) - ORDER.indexOf(b.code) || a.aluno.localeCompare(b.aluno, 'pt'));

const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9, legenda:10, verde:11,
  // linha pintada de tinta laranja: cota parcelada
  txtP:12, dataP:13, moedaP:14, numP:15 };
// O segundo argumento pinta a célula. Uso para a linha inteira de quem dividiu
// a cota em mais de uma parcela — é a leitura que a unidade faz de relance.
const c = (v, p) => ({ v: v == null ? '' : String(v), t:'s', s: p ? S.txtP : S.txt });
const m = (v, p) => v == null ? { v:'', t:'s', s: p ? S.moedaP : S.moeda } : { v, t:'m', s: p ? S.moedaP : S.moeda };
const n_ = (v, p) => v == null ? { v:'', t:'s', s: p ? S.numP : S.num } : { v, t:'n', s: p ? S.numP : S.num };
const dt = (v, p) => v ? { v, t:'d', s: p ? S.dataP : S.data } : { v:'', t:'s', s: p ? S.dataP : S.data };


/* ---- aba 1: resumo por unidade ---- */
function abaResumo() {
  const L2 = [];
  const t = (v, st) => ({ v, t:'s', s: st || S.txt });
  L2.push({ cells:[{ v:TITULO, t:'s', s:S.titulo }], h:24 });
  L2.push({ cells:[{ v:'Matrículas 2027 · fonte: '+path.basename(pFin), t:'s', s:S.legenda }] });
  L2.push({ cells:[{ v:(SO_RESERVA
    ? 'Só as parcelas cujo serviço é Reserva de vaga (EF1, EF2, EM, PM e PV). A 1ª Cota de '+
      'Mensalidade e as mensalidades ficam de fora. '
    : '')+
    'Baixada = valor baixado maior que zero com data de baixa válida. '+
    'A coluna DATA_BAIXA traz lixo numérico na maior parte das linhas, então os dois sinais têm de concordar. '+
    'Parcelas repetidas no export (mesmo REF_FINANCEIRO com dois status de boleto) contam uma vez só.',
    t:'s', s:S.legenda }], h:54 });
  L2.push({ cells:[] });
  L2.push({ cells:[t('Unidade',S.hdr), t('Alunos',S.hdr), t('Parcelas',S.hdr),
                   t('Valor baixado',S.hdr), t('Matrícula',S.hdr), t('Rematrícula',S.hdr)], h:26 });
  const por = {};
  alunos.forEach(a => {
    const u = por[a.code] = por[a.code] || { alunos:0, parc:0, valor:0, mat:0, rem:0 };
    u.alunos++; u.parc += a.n; u.valor += a.total;
    if (a.tipo === 'REMATRÍCULA') u.rem++; else u.mat++;
  });
  ORDER.filter(u => por[u]).forEach(u => {
    const v = por[u];
    L2.push({ cells:[c(CODE[u]), n_(v.alunos), n_(v.parc), m(v.valor), n_(v.mat), n_(v.rem)] });
  });
  const T = Object.values(por).reduce((a, v) => ({
    alunos:a.alunos+v.alunos, parc:a.parc+v.parc, valor:a.valor+v.valor,
    mat:a.mat+v.mat, rem:a.rem+v.rem }), { alunos:0, parc:0, valor:0, mat:0, rem:0 });
  L2.push({ cells:[t('Rede',S.hdr), n_(T.alunos), n_(T.parc), m(T.valor), n_(T.mat), n_(T.rem)] });

  // Quem lancou cada matricula. Nao e quem operou a baixa - o export do Totvs
  // nao traz esse campo -, e sim o usuario dono do registro no financeiro.
  L2.push({ cells:[] });
  const nP = alunos.filter(a => a.parcelado).length;
  L2.push({ cells:[t('Cota parcelada', S.titulo)], h:22 });
  L2.push({ cells:[t('As linhas em laranja, nas abas Por aluno e Parcelas, são de quem dividiu a '+
    'cota de entrada em mais de uma cobrança — o plano do Totvs prevê duas ou mais. '+
    'Vale para a Reserva de vaga: no serviço 1ª Cota de Mensalidade o plano de 12 parcelas é a '+
    'anuidade, e a cota continua sendo uma parcela só. '+
    'Parcelamento no cartão pelo Marketplace não conta: ali a escola recebe integral. '+
    (nP ? nP + ' de ' + alunos.length + ' alunos.' : 'Nenhum aluno nesta base.'),
    S.legenda)], h:42 });

  L2.push({ cells:[] });
  L2.push({ cells:[t('Quem lançou no Totvs', S.titulo)], h:22 });
  L2.push({ cells:[t('Usuário responsável pelo registro da matrícula. O relatório do Totvs não '+
    'traz o operador da baixa; peça a coluna de usuário da baixa para incluí-la aqui.',
    S.legenda)], h:30 });
  L2.push({ cells:[t('Quem lançou',S.hdr), t('Login',S.hdr), t('Alunos',S.hdr),
                   t('Parcelas',S.hdr), t('Valor baixado',S.hdr)], h:26 });
  const pu = {};
  baixadas.forEach(p => {
    const k = p.login || '(vazio)';
    const u = pu[k] = pu[k] || { nome: p.usuario || '(vazio)', ras: new Set(), parc:0, valor:0 };
    u.ras.add(p.ra); u.parc++; u.valor += p.baixado;
  });
  Object.entries(pu)
    .sort((a,b) => b[1].parc - a[1].parc || a[1].nome.localeCompare(b[1].nome,'pt'))
    .forEach(([login,u]) => L2.push({ cells:[
      c(u.nome), c(login), n_(u.ras.size), n_(u.parc), m(u.valor)] }));

  // Reserva de vaga: por onde entrou o dinheiro e se veio parcelado.
  if (COM_CANAL) {
    const res = SO_RESERVA ? baixadas : baixadas.filter(p => p.reserva);
    const esp = f => /RECREIO/i.test(String(f||'')) ? 600 : 300;
    const alu = new Map();
    res.forEach(p => {
      const a = alu.get(p.ra) || { canal:p.canal, plano:p.plano, filial:p.filial, parc:0, valor:0 };
      a.parc++; a.valor += p.baixado; alu.set(p.ra, a);
    });

    L2.push({ cells:[] });
    L2.push({ cells:[t('Reservas de vaga: por onde o dinheiro entrou', S.titulo)], h:22 });
    L2.push({ cells:[t('Marketplace = a transação está no export da Layers com status Pago. '+
      'Portal do aluno = a baixa existe no Totvs e não há transação paga correspondente no Marketplace.',
      S.legenda)], h:30 });
    L2.push({ cells:[t('Canal',S.hdr), t('Alunos',S.hdr), t('Parcelas',S.hdr), t('Valor baixado',S.hdr)], h:26 });
    const pc = {};
    [...alu.values()].forEach(a => { const o = pc[a.canal] = pc[a.canal] || { n:0, parc:0, valor:0 };
      o.n++; o.parc += a.parc; o.valor += a.valor; });
    Object.entries(pc).sort((a,b) => b[1].n - a[1].n)
      .forEach(([k,o]) => L2.push({ cells:[c(k), n_(o.n), n_(o.parc), m(o.valor)] }));
    const TR = Object.values(pc).reduce((a,o) => ({ n:a.n+o.n, parc:a.parc+o.parc, valor:a.valor+o.valor }),
      { n:0, parc:0, valor:0 });
    L2.push({ cells:[t('Total',S.hdr), n_(TR.n), n_(TR.parc), m(TR.valor)] });

    L2.push({ cells:[] });
    L2.push({ cells:[t('Reservas pagas pelo Portal: foram parceladas?', S.titulo)], h:22 });
    L2.push({ cells:[t('O plano vem da coluna DESCRICAO1 do Totvs. Parcela única = a reserva '+
      'foi cobrada de uma vez. Integral = o baixado cobre a cota da unidade (R$ 300, ou R$ 600 em Américas).',
      S.legenda)], h:30 });
    L2.push({ cells:[t('Plano de pagamento',S.hdr), t('Alunos',S.hdr), t('Integral',S.hdr),
                     t('Parcial',S.hdr), t('Valor baixado',S.hdr)], h:26 });
    const pp = {};
    [...alu.values()].filter(a => /^Portal/.test(a.canal)).forEach(a => {
      const k = a.plano || '(sem plano)';
      const o = pp[k] = pp[k] || { n:0, ok:0, par:0, valor:0 };
      o.n++; o.valor += a.valor;
      if (a.valor >= esp(a.filial) * 0.98) o.ok++; else o.par++;
    });
    Object.entries(pp).sort((a,b) => b[1].n - a[1].n)
      .forEach(([k,o]) => L2.push({ cells:[c(k), n_(o.n), n_(o.ok), n_(o.par), m(o.valor)] }));
  }

  return { nome:'Resumo', cols:[30,20,11,16,14,12], linhas:L2 };
}

/* ---- aba 2: um aluno por linha ---- */
const COLS_A = [
  ['Unidade',20],['Sigla',7],['Aluno',34],['RA',14],['Tipo',13],
  ['Situação no Totvs',26],['Curso',20],['Série',12],['Turma / trilha',30],
  ['Parcelas baixadas',10],['Valor baixado',14],['1ª baixa',12],['Última baixa',12],
  ['Canal do pagamento',30],['Como pagou a cota',26],['Quem lançou no Totvs',26],
  ['Serviços',34],['Responsável financeiro',30]
];
function abaAlunos() {
  const L2 = [{ cells: COLS_A.map(h => ({ v:h[0], t:'s', s:S.hdr })), h:30 }];
  alunos.forEach(a => { const p = a.parcelado; L2.push({ cells:[
    c(CODE[a.code],p), c(a.code,p), c(a.aluno,p), c(a.ra,p), c(a.tipo,p),
    c(a.situacao,p), c(a.curso,p), c(a.serie,p), c(a.descricao,p),
    n_(a.n,p), m(a.total,p), dt(a.primeira,p), dt(a.ultima,p), c(a.canal,p),
    c(p ? [...a.planos].join(' · ') : 'Parcela única', p),
    c([...a.usuarios].join(' · '),p), c([...a.servicos].join(' · '),p), c(a.resp,p)
  ], h:20 }); });
  const ult = colName(COLS_A.length-1) + L2.length;
  return { nome:'Por aluno', cols:COLS_A.map(h=>h[1]), linhas:L2, congelar:1, filtro:'A1:'+ult };
}

/* ---- aba 3: uma parcela por linha ---- */
const COLS_P = [
  ['Unidade',20],['Sigla',7],['Aluno',34],['RA',14],['Tipo',13],
  ['Situação no Totvs',26],['Serviço',30],['Plano de pagamento',28],['Parcela',9],['Cota',7],
  ['Valor original',13],['Bolsa / deduções',14],['Valor baixado',13],
  ['Vencimento',12],['Data da baixa',12],['Canal do pagamento',30],['Quem lançou no Totvs',26],['Login',20],
  ['Status do boleto',15],['Situação do contrato',18]
];
function abaParcelas() {
  const L2 = [{ cells: COLS_P.map(h => ({ v:h[0], t:'s', s:S.hdr })), h:30 }];
  baixadas.forEach(p => { const z = p.parcelado; L2.push({ cells:[
    c(CODE[p.code],z), c(p.code,z), c(p.aluno,z), c(p.ra,z), c(p.tipo,z),
    c(p.situacao,z), c(p.servico,z), c(p.plano,z), c(p.parcela,z), c(p.cota,z),
    m(p.original,z), m(p.deducoes,z), m(p.baixado,z),
    dt(p.venc,z), dt(p.dtBaixa,z), c(p.canal,z), c(p.usuario,z), c(p.login,z),
    c(p.boleto,z), c(p.contrato,z)
  ], h:20 }); });
  const ult = colName(COLS_P.length-1) + L2.length;
  return { nome:'Parcelas', cols:COLS_P.map(h=>h[1]), linhas:L2, congelar:1, filtro:'A1:'+ult };
}

const arq = path.join(SAIDA, ARQUIVO);
fs.writeFileSync(arq, pasta([abaResumo(), abaAlunos(), abaParcelas()]));

console.log((SO_RESERVA ? 'reservas de vaga baixadas:' : 'parcelas baixadas:'),
  baixadas.length, '· alunos:', alunos.length);
console.log('valor total: R$', baixadas.reduce((a,p)=>a+p.baixado,0)
  .toLocaleString('pt-BR', { minimumFractionDigits:2, maximumFractionDigits:2 }));
const por = {}; alunos.forEach(a => por[a.code] = (por[a.code]||0)+1);
console.log('alunos por unidade:', ORDER.filter(u=>por[u]).map(u=>u+':'+por[u]).join('  '));
console.log('->', arq, Math.round(fs.statSync(arq).size/1024)+' KB');
