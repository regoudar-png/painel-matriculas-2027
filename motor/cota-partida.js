/* Alunos de captação cuja COTA DE ENTRADA foi partida em duas parcelas —
 * R$ 300 nas unidades, R$ 600 no Recreio — com uma parcela baixada e a outra
 * em aberto.
 *
 *   node motor/cota-partida.js [pasta-de-saida]
 *
 * Não confundir com a 1ª cota do plano anual (R$ 975 a R$ 3.097): essa é a
 * entrada da anuidade, outro assunto. Aqui é a cota do bolsão dividida em duas,
 * pelos planos "RESERVA DE VAGA (2x NO BOLETO)" e equivalentes.
 *
 * A base de 03/09 não traz baixa; quem diz o que foi pago é a Ficha Financeira
 * de 02/09, casada parcela a parcela pelo IDLAN. */

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
const TESTE = /\bteste?s?\b/i;
const HOJE = '2026-09-03';
const ENTRADA = /Reserva de vaga|1ª Cota de Mensalidade/i;
// Cota de entrada esperada: R$ 600 no Recreio, R$ 300 nas outras unidades.
const cotaCheia = u => u === 'AM' ? 600 : 300;

/* ---------------- baixas, parcela a parcela ---------------- */
const baixa = new Map();
{
  const abas = lerPlanilha(P_FICHA).abas;
  const L = abas.filter(a => a.linhas).sort((a, b) => b.linhas.length - a.linhas.length)[0].linhas;
  L.slice(1).filter(r => r && r[35]).forEach(r => {
    const k = String(r[35]).trim();
    const v = num(r[44]), d = ehData(r[45]) ? String(r[45]).trim() : null;
    const atual = baixa.get(k);
    if (!atual || (v > 0 && d && !(atual.valor > 0 && atual.data)))
      baixa.set(k, { valor: v, data: d, boleto: r[47] || null, contrato: r[26] || null });
  });
}

/* ---------------- base de 03/09 ---------------- */
const L = lerPlanilha(P_BASE).abas[0].linhas;
const H = L[0], col = n => H.indexOf(n);
const C = {
  fil: col('FILIAL'), ra: col('RA'), aluno: col('ALUNO'), tipo: col('TIPO MATRICULA'),
  status: col('STATUS'), curso: col('CURSO'), serie: col('SERIE'), desc: col('DESCRICAO'),
  serv: col('SERVICO'), parcela: col('PARCELA'), plano: col('PLANO DE PAGAMENTO'),
  idlan: col('IDLAN'), comp: col('COMPETENCIA'), venc: col('DATA VENCIMENTO'),
  orig: col('VALOR ORIGINAL'), bolsa: col('BOLSA R$'), pct: col('BOLSA %'),
  liq: col('VALOR LIQUIDO (ORIGINAL - BOLSA)')
};

// Uma cota partida é o mesmo serviço de entrada aparecendo mais de uma vez para
// o mesmo aluno. É isso que distingue da 1ª cota do plano anual, que aparece
// uma vez só e com valor de mensalidade.
const grupos = new Map();
L.slice(1).filter(r => r && r[C.ra] && r[C.tipo] === 'MATRÍCULA'
    && ENTRADA.test(String(r[C.serv])) && !TESTE.test(String(r[C.aluno] || '')))
  .forEach(r => {
    const k = r[C.ra] + '|' + String(r[C.serv]);
    if (!grupos.has(k)) grupos.set(k, {
      ra: r[C.ra], aluno: String(r[C.aluno] || '').trim(), u: unidade(r[C.fil]),
      status: r[C.status], curso: r[C.curso], serie: r[C.serie], trilha: r[C.desc],
      servico: r[C.serv], plano: r[C.plano], linhas: []
    });
    grupos.get(k).linhas.push(r);
  });

const bx = r => baixa.get(String(r[C.idlan] || '').trim()) || { valor: 0, data: null, boleto: null, contrato: null };

const casos = [...grupos.values()].filter(g => g.linhas.length > 1);
casos.forEach(g => {
  g.linhas.sort((a, b) => num(a[C.parcela]) - num(b[C.parcela]) ||
    ((paraISO(a[C.venc]) || '') < (paraISO(b[C.venc]) || '') ? -1 : 1));
  g.pagas = g.linhas.filter(r => bx(r).valor > 0);
  g.abertas = g.linhas.filter(r => !(bx(r).valor > 0));
  g.cobrado = g.linhas.reduce((s, r) => s + num(r[C.liq]), 0);
  g.pago = g.pagas.reduce((s, r) => s + bx(r).valor, 0);
  g.aberto = g.abertas.reduce((s, r) => s + num(r[C.liq]), 0);
  g.cheia = cotaCheia(g.u);
  g.vencAberto = g.abertas.map(r => r[C.venc])
    .sort((x, y) => (paraISO(x) || '') < (paraISO(y) || '') ? -1 : 1)[0] || null;
  g.vencido = g.vencAberto ? paraISO(g.vencAberto) < HOJE : false;
});

const alvo = casos.filter(g => g.pagas.length >= 1 && g.abertas.length >= 1)
  .sort((a, b) => ORDER.indexOf(a.u) - ORDER.indexOf(b.u) || a.aluno.localeCompare(b.aluno, 'pt'));

/* ---------------- planilha ---------------- */
const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9, legenda:10,
  verde:11, txtP:12, dataP:13, moedaP:14, numP:15 };
const c = (v, p) => ({ v: v == null ? '' : String(v), t:'s', s: p ? S.txtP : S.txt });
const m = (v, p) => v == null ? { v:'', t:'s', s: p ? S.moedaP : S.moeda } : { v, t:'m', s: p ? S.moedaP : S.moeda };
const n_ = (v, p) => v == null ? { v:'', t:'s', s: p ? S.numP : S.num } : { v, t:'n', s: p ? S.numP : S.num };
const dt = (v, p) => { const i = paraISO(v);
  return i ? { v:i, t:'d', s: p ? S.dataP : S.data } : { v:'', t:'s', s: p ? S.dataP : S.data }; };
const t = (v, st) => ({ v: v == null ? '' : String(v), t:'s', s: st || S.txt });
const brl = v => 'R$ ' + Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* ---- aba 1: resumo ---- */
function abaResumo() {
  const L2 = [];
  L2.push({ cells:[t('Cota de entrada partida em duas parcelas — captação 2027', S.titulo)], h:24 });
  L2.push({ cells:[t('Base: ' + path.basename(P_BASE) + ' · baixas conferidas em ' +
    path.basename(P_FICHA) + ' · posição de 3 de setembro de 2026', S.legenda)] });
  L2.push({ cells:[] });
  L2.push({ cells:[t('A cota de entrada do bolsão — R$ 300 nas unidades, R$ 600 no Recreio — foi '+
    'dividida em duas parcelas. A primeira foi paga e a segunda está em aberto. Não confundir '+
    'com a 1ª cota do plano anual, de R$ 975 a R$ 3.097, que é a entrada da anuidade e aparece '+
    'uma vez só.', S.legenda)], h:54 });
  L2.push({ cells:[] });

  L2.push({ cells:[t('',S.hdr), t('Alunos',S.hdr), t('Cobrado',S.hdr), t('Pago',S.hdr),
    t('Em aberto',S.hdr)], h:26 });
  const bloco = (rot, l) => L2.push({ cells:[c(rot), n_(l.length),
    m(l.reduce((s,g)=>s+g.cobrado,0)), m(l.reduce((s,g)=>s+g.pago,0)),
    m(l.reduce((s,g)=>s+g.aberto,0))] });
  bloco('Recreio — cota de R$ 600 em 2x', alvo.filter(g => g.u === 'AM'));
  bloco('Unidades — cota de R$ 300 em 2x', alvo.filter(g => g.u !== 'AM'));
  L2.push({ cells:[t('Total',S.hdr), n_(alvo.length),
    m(alvo.reduce((s,g)=>s+g.cobrado,0)), m(alvo.reduce((s,g)=>s+g.pago,0)),
    m(alvo.reduce((s,g)=>s+g.aberto,0))] });
  L2.push({ cells:[] });

  L2.push({ cells:[t('Unidade',S.hdr), t('Alunos',S.hdr), t('Cota cheia',S.hdr),
    t('Em aberto',S.hdr), t('Vencidas',S.hdr)], h:26 });
  const pu = {};
  alvo.forEach(g => { const o = pu[g.u] = pu[g.u] || { n:0, ab:0, vc:0 };
    o.n++; o.ab += g.aberto; if (g.vencido) o.vc++; });
  ORDER.filter(u => pu[u]).forEach(u => L2.push({ cells:[c(CODE[u]), n_(pu[u].n),
    m(cotaCheia(u)), m(pu[u].ab), pu[u].vc ? n_(pu[u].vc) : c('—')] }));
  L2.push({ cells:[] });
  L2.push({ cells:[t('Como identifiquei', S.titulo)], h:22 });
  L2.push({ cells:[t('O mesmo serviço de entrada aparece mais de uma vez para o mesmo aluno na '+
    'base de 03/09 — é o plano "RESERVA DE VAGA (2x NO BOLETO)" e equivalentes. Cada parcela tem '+
    'seu próprio IDLAN, que casa com o REF_FINANCEIRO da ficha financeira: por isso dá para dizer '+
    'qual das duas foi baixada. Varri as ' + [...grupos.values()].length + ' cobranças de entrada '+
    'da captação e encontrei ' + casos.length + ' cotas partidas — todas com uma parcela paga e '+
    'uma em aberto, nenhuma com as duas pagas ou as duas abertas.', S.legenda)], h:66 });
  return { nome:'Resumo', cols:[38,10,15,15,15], linhas:L2 };
}

/* ---- aba 2: um aluno por linha ---- */
const COLS_A = [
  ['Unidade',20],['Sigla',7],['Aluno',34],['RA',13],['Situação no Totvs',26],
  ['Curso',20],['Série',10],['Turma / trilha',30],['Serviço',30],['Plano de pagamento',34],
  ['Cota cheia',12],['1ª parcela',12],['Baixada em',13],['2ª parcela',12],['Vence em',13],
  ['Em aberto',12],['Prazo',10]
];
function abaAlunos() {
  const L2 = [{ cells: COLS_A.map(h => t(h[0], S.hdr)), h:30 }];
  alvo.forEach(g => {
    const p = g.vencido;
    const paga = g.pagas[0], abre = g.abertas[0];
    L2.push({ cells:[
      c(CODE[g.u], p), c(g.u, p), c(g.aluno, p), c(g.ra, p), c(g.status, p),
      c(g.curso, p), c(g.serie, p), c(g.trilha, p), c(g.servico, p), c(g.plano, p),
      m(g.cheia, p), m(num(paga[C.liq]), p), dt(bx(paga).data, p),
      m(num(abre[C.liq]), p), dt(abre[C.venc], p),
      m(g.aberto, p), c(p ? 'VENCIDA' : 'a vencer', p)
    ], h:20 });
  });
  const ult = colName(COLS_A.length - 1) + L2.length;
  return { nome:'Por aluno', cols:COLS_A.map(h => h[1]), linhas:L2, congelar:1, filtro:'A1:' + ult };
}

/* ---- aba 3: as duas linhas de cada aluno, cruas ---- */
const COLS_L = [
  ['Unidade',20],['Sigla',7],['Aluno',34],['RA',13],['Situação no Totvs',26],
  ['Serviço',30],['Plano de pagamento',34],['IDLAN',11],['Parcela',8],['Competência',12],
  ['Vencimento',12],['Valor original',13],['Bolsa R$',11],['Bolsa %',9],['Valor líquido',13],
  ['Baixado',13],['Data da baixa',13],['Status do boleto',16],['Situação do contrato',18],['Situação',13]
];
function abaLinhas() {
  const L2 = [{ cells: COLS_L.map(h => t(h[0], S.hdr)), h:30 }];
  alvo.forEach(g => g.linhas.forEach(r => {
    const b = bx(r), aberta = !(b.valor > 0);
    L2.push({ cells:[
      c(CODE[g.u], aberta), c(g.u, aberta), c(g.aluno, aberta), c(g.ra, aberta),
      c(g.status, aberta), c(r[C.serv], aberta), c(r[C.plano], aberta),
      c(r[C.idlan], aberta), c(r[C.parcela], aberta), c(r[C.comp], aberta),
      dt(r[C.venc], aberta), m(num(r[C.orig]), aberta), m(num(r[C.bolsa]), aberta),
      n_(num(r[C.pct]), aberta), m(num(r[C.liq]), aberta),
      b.valor > 0 ? m(b.valor) : m(null, true), dt(b.data, aberta),
      c(b.boleto, aberta), c(b.contrato, aberta),
      c(aberta ? 'EM ABERTO' : 'baixada', aberta)
    ], h:20 });
  }));
  const ult = colName(COLS_L.length - 1) + L2.length;
  return { nome:'Parcelas', cols:COLS_L.map(h => h[1]), linhas:L2, congelar:1, filtro:'A1:' + ult };
}

const arq = path.join(SAIDA, 'Cota de entrada partida - captação 2027.xlsx');
fs.writeFileSync(arq, pasta([abaResumo(), abaAlunos(), abaLinhas()]));

console.log('cobranças de entrada varridas:', [...grupos.values()].reduce((s,g)=>s+g.linhas.length,0),
  'em', [...grupos.values()].length, 'grupos');
console.log('cotas partidas em 2+ parcelas:', casos.length,
  '| com uma baixada e uma em aberto:', alvo.length);
console.log('cobrado:', brl(alvo.reduce((s,g)=>s+g.cobrado,0)),
  '| pago:', brl(alvo.reduce((s,g)=>s+g.pago,0)),
  '| em aberto:', brl(alvo.reduce((s,g)=>s+g.aberto,0)));
alvo.forEach(g => console.log('  ', g.u, g.aluno.slice(0,32).padEnd(34), g.status.slice(0,26).padEnd(28),
  'cota', brl(g.cheia).padStart(10), '| pagou', brl(g.pago).padStart(10),
  '| deve', brl(g.aberto).padStart(10), '| vence', g.vencAberto, g.vencido ? 'VENCIDA' : ''));
console.log('->', arq, Math.round(fs.statSync(arq).size / 1024) + ' KB');
