/* Lista dos alunos de captação com duas cobranças de entrada.
 *
 *   node motor/lista-duas-cotas.js [pasta-de-saida]
 *
 * Três abas: o diagnóstico da causa, uma linha por aluno e — o que foi pedido —
 * as linhas cruas das duas cobranças, a reserva baixada e a 1ª cota gerada pelo
 * contrato. A baixa vem da ficha financeira de 02/09, casada linha a linha pelo
 * IDLAN, que é o mesmo identificador do REF_FINANCEIRO da ficha. */

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');
const { pasta, colName } = require('./escrever-xlsx.js');

const B = 'C:/Users/ti/Desktop/Matrículas 2027/';
const P_BASE = B + '03 de setembro/Totovs_Base financeira.XLSX';
const P_FICHA = B + '02 de setembro/Ficha Financeira.XLSX';
const SAIDA = process.argv[2] || B;

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

const num = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; };
const ehData = v => /^\d{2}\/\d{2}\/\d{4}$/.test(String(v || '').trim());
const paraISO = v => ehData(v) ? String(v).trim().split('/').reverse().join('-') : null;
const serial = v => { const x = num(v);
  return x > 20000 ? new Date(Date.UTC(1899, 11, 30) + x * 864e5).toISOString().slice(0, 10) : null; };
const TESTE = /\bteste?s?\b/i;
const HOJE = '2026-09-03';

const RESERVA = /Reserva de vaga/i;
const COTA = /1ª Cota de Mensalidade/i;
const MENS = /(^|- )Mensalidade$/i;

/* ---------------- baixas, por IDLAN ---------------- */
const baixa = new Map();   // IDLAN -> {valor, data, boleto, situacao}
{
  const abas = lerPlanilha(P_FICHA).abas;
  const L = abas.filter(a => a.linhas).sort((a, b) => b.linhas.length - a.linhas.length)[0].linhas;
  L.slice(1).filter(r => r && r[35]).forEach(r => {
    const k = String(r[35]).trim();
    const v = num(r[44]), d = ehData(r[45]) ? String(r[45]).trim() : null;
    const atual = baixa.get(k);
    // A mesma parcela vem duas vezes quando o boleto tem dois status de
    // remessa: fico com a linha que registra a baixa.
    if (!atual || (v > 0 && d && !(atual.valor > 0 && atual.data)))
      baixa.set(k, { valor: v, data: d, boleto: r[47] || null, situacao: r[26] || null });
  });
}

/* ---------------- base de 03/09 ---------------- */
const L = lerPlanilha(P_BASE).abas[0].linhas;
const H = L[0], col = n => H.indexOf(n);
const C = {
  fil: col('FILIAL'), ra: col('RA'), aluno: col('ALUNO'), tipo: col('TIPO MATRICULA'),
  dtMat: col('DATA MATRICULA'), status: col('STATUS'), curso: col('CURSO'), serie: col('SERIE'),
  grade: col('CODGRADE'), desc: col('DESCRICAO'), turma: col('CODTURMA'), serv: col('SERVICO'),
  parcela: col('PARCELA'), plano: col('PLANO DE PAGAMENTO'), idlan: col('IDLAN'),
  comp: col('COMPETENCIA'), venc: col('DATA VENCIMENTO'), orig: col('VALOR ORIGINAL'),
  bolsa: col('BOLSA R$'), pct: col('BOLSA %'), descontos: col('DESCONTOS'),
  liq: col('VALOR LIQUIDO (ORIGINAL - BOLSA)'), cotaTab: col('1ª COTA TABELA'),
  mensTab: col('MENSALIDADE TABELA'), anuTab: col('ANUIDADE TABELA')
};

const cap = L.slice(1).filter(r => r && r[C.ra] && r[C.tipo] === 'MATRÍCULA'
  && !TESTE.test(String(r[C.aluno] || '')));

const alunos = new Map();
cap.forEach(r => {
  const ra = r[C.ra];
  if (!alunos.has(ra)) alunos.set(ra, {
    ra, aluno: String(r[C.aluno] || '').trim(), u: unidade(r[C.fil]), filial: r[C.fil],
    dtMat: serial(r[C.dtMat]), status: r[C.status], curso: r[C.curso], serie: r[C.serie],
    grade: r[C.grade], trilha: r[C.desc], turma: r[C.turma],
    res: [], cota: [], mens: []
  });
  const a = alunos.get(ra);
  const s = String(r[C.serv] || '');
  if (RESERVA.test(s)) a.res.push(r);
  else if (COTA.test(s)) a.cota.push(r);
  else if (MENS.test(s)) a.mens.push(r);
});

const todos = [...alunos.values()];
const bx = r => baixa.get(String(r[C.idlan] || '').trim()) || { valor: 0, data: null, boleto: null, situacao: null };
const pago = rs => rs.reduce((s, r) => s + (bx(r).valor || 0), 0);

const dupla = todos.filter(a => a.res.length && a.cota.length);
dupla.forEach(a => {
  a.pagoRes = pago(a.res); a.pagoCota = pago(a.cota);
  a.dtRes = a.res.map(r => bx(r).data).find(Boolean) || null;
  a.dtCota = a.cota.map(r => bx(r).data).find(Boolean) || null;
  a.abertoRes = a.res.reduce((s, r) => s + (bx(r).valor > 0 ? 0 : num(r[C.liq])), 0);
  a.abertoCota = a.cota.reduce((s, r) => s + (bx(r).valor > 0 ? 0 : num(r[C.liq])), 0);
  a.aberto = a.abertoRes + a.abertoCota;
  a.vencAberto = a.cota.concat(a.res).filter(r => !(bx(r).valor > 0))
    .map(r => r[C.venc]).sort((x, y) => paraISO(x) < paraISO(y) ? -1 : 1)[0] || null;
  a.vencido = a.vencAberto ? paraISO(a.vencAberto) < HOJE : false;
  a.mensLiq = a.mens.length ? num(a.mens[0][C.liq]) : null;
  a.pctBolsa = a.mens.length ? num(a.mens[0][C.pct]) : (a.cota[0] ? num(a.cota[0][C.pct]) : null);
  a.planoAnual = a.mens.length ? a.mens[0][C.plano] : (a.cota[0] ? a.cota[0][C.plano] : null);
});
dupla.sort((x, y) => ORDER.indexOf(x.u) - ORDER.indexOf(y.u) ||
  (y.vencido - x.vencido) || (y.aberto - x.aberto));

/* grupos, para o diagnóstico */
const soCota = todos.filter(a => !a.res.length && a.cota.length);
const soRes = todos.filter(a => a.res.length && !a.cota.length);
const certos = soRes.filter(a => a.mens.length);
const semPlano = soRes.filter(a => !a.mens.length);

/* ---------------- planilha ---------------- */
const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9, legenda:10,
  verde:11, txtP:12, dataP:13, moedaP:14, numP:15 };
const c = (v, p) => ({ v: v == null ? '' : String(v), t:'s', s: p ? S.txtP : S.txt });
const m = (v, p) => v == null ? { v:'', t:'s', s: p ? S.moedaP : S.moeda } : { v, t:'m', s: p ? S.moedaP : S.moeda };
const n_ = (v, p) => v == null ? { v:'', t:'s', s: p ? S.numP : S.num } : { v, t:'n', s: p ? S.numP : S.num };
const dt = (v, p) => { const i = paraISO(v);
  return i ? { v:i, t:'d', s: p ? S.dataP : S.data } : { v:'', t:'s', s: p ? S.dataP : S.data }; };
const t = (v, st) => ({ v: v == null ? '' : String(v), t:'s', s: st || S.txt });

/* ---- aba 1: diagnóstico ---- */
function abaDiag() {
  const L2 = [];
  const total = todos.length;
  L2.push({ cells:[t('Duas cobranças de entrada na captação 2027', S.titulo)], h:24 });
  L2.push({ cells:[t('Base: ' + path.basename(P_BASE) + ' · baixas conferidas em ' +
    path.basename(P_FICHA) + ' · posição de 3 de setembro de 2026', S.legenda)] });
  L2.push({ cells:[] });

  L2.push({ cells:[t('O que aconteceu', S.titulo)], h:22 });
  L2.push({ cells:[t('Existem dois contratos por aluno. O do bolsão cobra a Reserva de vaga (R$ 300). '+
    'Ao formalizar a matrícula, a unidade escolhe o plano de pagamento do ano — e é aí que os '+
    'caminhos se separam. O plano "01 COTA C/30% DE DESC. + MENSALIDADE" gera uma 1ª cota nova, '+
    'ignorando a reserva já paga. O plano "01 COTA + 12 PARCELAS" lança só as mensalidades, '+
    'tratando a reserva como a entrada. Quem usou o primeiro ficou com duas cobranças.',
    S.legenda)], h:66 });
  L2.push({ cells:[] });

  L2.push({ cells:[t('Grupo', S.hdr), t('Alunos', S.hdr), t('Matrícula feita entre', S.hdr),
    t('Plano do ano', S.hdr), t('Situação', S.hdr)], h:26 });
  const faixa = l => { const d = l.map(a => a.dtMat).filter(Boolean).sort();
    return d.length ? d[0] + ' a ' + d[d.length - 1] : '—'; };
  L2.push({ cells:[c('Reserva paga + 1ª cota gerada'), n_(dupla.length), c(faixa(dupla)),
    c('01 COTA C/30% DE DESC. + MENSALIDADE'), c('cobrado duas vezes')] });
  L2.push({ cells:[c('Reserva paga, sem 1ª cota'), n_(certos.length), c(faixa(certos)),
    c('01 COTA + 12 PARCELAS'), c('correto — a reserva fez o papel da cota')] });
  L2.push({ cells:[c('Reserva paga, plano do ano não gerado'), n_(semPlano.length), c(faixa(semPlano)),
    c('—'), c('risco: pode repetir o erro quando o plano for lançado')] });
  L2.push({ cells:[c('Sem reserva, só a 1ª cota'), n_(soCota.length), c(faixa(soCota)),
    c('vários'), c('correto — entraram depois de 19/08, sem reserva')] });
  L2.push({ cells:[t('Total na captação', S.hdr), n_(total), c(''), c(''), c('')] });
  L2.push({ cells:[] });

  L2.push({ cells:[t('O corte é 19 de agosto', S.titulo)], h:22 });
  L2.push({ cells:[t('Toda matrícula feita de 8 a 18 de agosto tem Reserva de vaga: 130 de 130. '+
    'De 19 de agosto em diante, nenhuma tem: 0 de 140. O produto Reserva de vaga saiu de uso '+
    'entre o primeiro e o segundo bolsão, e o problema está inteiro na turma do primeiro.',
    S.legenda)], h:42 });
  L2.push({ cells:[] });

  const venc = dupla.filter(a => a.vencido);
  L2.push({ cells:[t('O dinheiro', S.hdr), t('Alunos', S.hdr), t('Valor', S.hdr)], h:26 });
  L2.push({ cells:[c('Em aberto, total'), n_(dupla.length), m(dupla.reduce((s, a) => s + a.aberto, 0))] });
  L2.push({ cells:[c('Já vencido'), n_(venc.length), m(venc.reduce((s, a) => s + a.aberto, 0))] });
  L2.push({ cells:[c('A vencer'), n_(dupla.length - venc.length),
    m(dupla.filter(a => !a.vencido).reduce((s, a) => s + a.aberto, 0))] });
  L2.push({ cells:[c('Reserva já paga (entrou no caixa)'), n_(dupla.filter(a => a.pagoRes > 0).length),
    m(dupla.reduce((s, a) => s + a.pagoRes, 0))] });
  L2.push({ cells:[] });

  L2.push({ cells:[t('Por unidade', S.hdr), t('Alunos', S.hdr), t('Em aberto', S.hdr),
    t('Vencido', S.hdr), t('Reserva paga', S.hdr)], h:26 });
  const pu = {};
  dupla.forEach(a => { const o = pu[a.u] = pu[a.u] || { n:0, ab:0, vc:0, rp:0 };
    o.n++; o.ab += a.aberto; if (a.vencido) o.vc += a.aberto; o.rp += a.pagoRes; });
  ORDER.filter(u => pu[u]).sort((x, y) => pu[y].ab - pu[x].ab).forEach(u => {
    const o = pu[u];
    L2.push({ cells:[c(CODE[u]), n_(o.n), m(o.ab), o.vc ? m(o.vc) : c('—'), m(o.rp)] });
  });
  const T = Object.values(pu).reduce((a, o) => ({ n:a.n+o.n, ab:a.ab+o.ab, vc:a.vc+o.vc, rp:a.rp+o.rp }),
    { n:0, ab:0, vc:0, rp:0 });
  L2.push({ cells:[t('Rede', S.hdr), n_(T.n), m(T.ab), m(T.vc), m(T.rp)] });

  return { nome:'Diagnóstico', cols:[38,10,22,40,44], linhas:L2 };
}

/* ---- aba 2: um aluno por linha ---- */
const COLS_A = [
  ['Unidade',20],['Sigla',7],['Aluno',36],['RA',13],['Situação no Totvs',26],
  ['Data da matrícula',14],['Curso',20],['Série',10],['Turma',8],['Turma / trilha',30],
  ['Reserva — valor',14],['Reserva — baixada em',15],['1ª cota — valor',14],
  ['1ª cota — vence em',15],['1ª cota — baixada em',15],
  ['Em aberto',13],['Prazo',10],['Mensalidade',13],['Bolsa %',9],['Plano do ano',44]
];
function abaAlunos() {
  const L2 = [{ cells: COLS_A.map(h => t(h[0], S.hdr)), h:30 }];
  dupla.forEach(a => {
    const p = a.vencido;
    const res = a.res[0], cot = a.cota[0];
    L2.push({ cells:[
      c(CODE[a.u], p), c(a.u, p), c(a.aluno, p), c(a.ra, p), c(a.status, p),
      dt(a.dtMat ? a.dtMat.split('-').reverse().join('/') : null, p),
      c(a.curso, p), c(a.serie, p), c(a.turma, p), c(a.trilha, p),
      m(res ? num(res[C.liq]) : null, p), dt(a.dtRes, p),
      m(cot ? num(cot[C.liq]) : null, p), dt(cot ? cot[C.venc] : null, p), dt(a.dtCota, p),
      m(a.aberto, p), c(p ? 'VENCIDA' : 'a vencer', p),
      m(a.mensLiq, p), n_(a.pctBolsa, p), c(a.planoAnual, p)
    ], h:20 });
  });
  const ult = colName(COLS_A.length - 1) + L2.length;
  return { nome:'Por aluno', cols:COLS_A.map(h => h[1]), linhas:L2, congelar:1, filtro:'A1:' + ult };
}

/* ---- aba 3: as linhas cruas das duas cobranças ---- */
const COLS_L = [
  ['Unidade',20],['Sigla',7],['Aluno',36],['RA',13],['Situação no Totvs',26],
  ['Cobrança',12],['Serviço',30],['Plano de pagamento',44],['IDLAN',11],
  ['Parcela',8],['Competência',12],['Vencimento',12],
  ['Valor original',13],['Bolsa R$',12],['Bolsa %',9],['Descontos',12],['Valor líquido',13],
  ['Baixado',13],['Data da baixa',13],['Status do boleto',16],['Situação do contrato',18],['Situação',14]
];
function abaLinhas() {
  const L2 = [{ cells: COLS_L.map(h => t(h[0], S.hdr)), h:30 }];
  dupla.forEach(a => {
    const linhas = a.res.map(r => ['Reserva', r]).concat(a.cota.map(r => ['1ª cota', r]));
    linhas.sort((x, y) => (paraISO(x[1][C.venc]) || '') < (paraISO(y[1][C.venc]) || '') ? -1 : 1);
    linhas.forEach(([rot, r]) => {
      const b = bx(r);
      const abertaAqui = !(b.valor > 0);
      L2.push({ cells:[
        c(CODE[a.u], abertaAqui), c(a.u, abertaAqui), c(a.aluno, abertaAqui), c(a.ra, abertaAqui),
        c(a.status, abertaAqui), c(rot, abertaAqui), c(r[C.serv], abertaAqui),
        c(r[C.plano], abertaAqui), c(r[C.idlan], abertaAqui), c(r[C.parcela], abertaAqui),
        c(r[C.comp], abertaAqui), dt(r[C.venc], abertaAqui),
        m(num(r[C.orig]), abertaAqui), m(num(r[C.bolsa]), abertaAqui),
        n_(num(r[C.pct]), abertaAqui), m(num(r[C.descontos]), abertaAqui),
        m(num(r[C.liq]), abertaAqui),
        b.valor > 0 ? m(b.valor) : m(null, true), dt(b.data, abertaAqui),
        c(b.boleto, abertaAqui), c(b.situacao, abertaAqui),
        c(abertaAqui ? 'EM ABERTO' : 'baixada', abertaAqui)
      ], h:20 });
    });
  });
  const ult = colName(COLS_L.length - 1) + L2.length;
  return { nome:'Cobranças', cols:COLS_L.map(h => h[1]), linhas:L2, congelar:1, filtro:'A1:' + ult };
}

const arq = path.join(SAIDA, 'Duas cotas na captação - 2027.xlsx');
fs.writeFileSync(arq, pasta([abaDiag(), abaAlunos(), abaLinhas()]));

const brl = v => 'R$ ' + Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
console.log('alunos com duas cobranças:', dupla.length,
  '| em aberto:', brl(dupla.reduce((s, a) => s + a.aberto, 0)));
console.log('linhas de cobrança:', dupla.reduce((s, a) => s + a.res.length + a.cota.length, 0));
console.log('grupos: dupla', dupla.length, '| certos', certos.length,
  '| sem plano do ano', semPlano.length, '| sem reserva', soCota.length);
console.log('->', arq, Math.round(fs.statSync(arq).size / 1024) + ' KB');
