/* Rastreio da base financeira de 03/09.
 *
 *   node motor/rastreio-0309.js
 *
 * Duas perguntas:
 *   1. quem, na captação, está com duas cobranças de entrada — a reserva de
 *      vaga baixada e a 1ª cota ainda em aberto;
 *   2. como está a matrícula por unidade e por status no resto da base.
 *
 * A base de 03/09 não traz coluna de baixa, então quem diz o que foi pago é a
 * Ficha Financeira de 02/09, cruzada por RA. */

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');

const B = 'C:/Users/ti/Desktop/Matrículas 2027/';
const P_NOVA = B + '03 de setembro/Totovs_Base financeira.XLSX';
const P_BAIXA = B + '02 de setembro/Ficha Financeira.XLSX';

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
const TESTE = /\bteste?s?\b/i;

const RESERVA = /Reserva de vaga/i;
const COTA = /1ª Cota de Mensalidade/i;
const MENSALIDADE = /(^|- )Mensalidade$/i;

/* ---------- o que já foi baixado, da ficha de 02/09 ---------- */
const baixaPorRA = new Map();   // RA -> {res, cota, dtRes, dtCota}
// IDLAN da base de 03/09 é o mesmo identificador do REF_FINANCEIRO da ficha,
// então dá para casar parcela com parcela em vez de somar por aluno.
const baixaPorLan = new Map();  // IDLAN -> {valor, data}
{
  const abas = lerPlanilha(P_BAIXA).abas;
  const L = abas.filter(a => a.linhas).sort((a, b) => b.linhas.length - a.linhas.length)[0].linhas;
  // A mesma parcela vem duas vezes quando o boleto tem dois status de remessa.
  const vistos = new Set();
  L.slice(1).filter(r => r && r[3] && num(r[44]) > 0 && ehData(r[45])).forEach(r => {
    const ref = String(r[35] || '').trim();
    if (ref) { if (vistos.has(ref)) return; vistos.add(ref); }
    const lan = String(r[35] || '').trim();
    if (lan) baixaPorLan.set(lan, { valor: num(r[44]), data: String(r[45]).trim() });
    const o = baixaPorRA.get(r[3]) || { res: 0, cota: 0, dtRes: null, dtCota: null };
    const serv = String(r[34] || '');
    if (RESERVA.test(serv)) { o.res += num(r[44]); o.dtRes = o.dtRes || String(r[45]).trim(); }
    else if (COTA.test(serv)) { o.cota += num(r[44]); o.dtCota = o.dtCota || String(r[45]).trim(); }
    baixaPorRA.set(r[3], o);
  });
}

/* ---------- base de 03/09 ---------- */
const L = lerPlanilha(P_NOVA).abas[0].linhas;
const H = L[0], col = n => H.indexOf(n);
const C = {
  fil: col('FILIAL'), ra: col('RA'), aluno: col('ALUNO'), tipo: col('TIPO MATRICULA'),
  status: col('STATUS'), curso: col('CURSO'), serie: col('SERIE'), desc: col('DESCRICAO'),
  serv: col('SERVICO'), parcela: col('PARCELA'), plano: col('PLANO DE PAGAMENTO'),
  venc: col('DATA VENCIMENTO'), orig: col('VALOR ORIGINAL'), bolsa: col('BOLSA R$'),
  pct: col('BOLSA %'), liq: col('VALOR LIQUIDO (ORIGINAL - BOLSA)'), idlan: col('IDLAN'),
  anuidade: col('ANUIDADE TABELA')
};

const linhas = L.slice(1).filter(r => r && r[C.ra] && !TESTE.test(String(r[C.aluno] || '')));

const alunos = new Map();
linhas.forEach(r => {
  const ra = r[C.ra];
  if (!alunos.has(ra)) alunos.set(ra, {
    ra, aluno: String(r[C.aluno] || '').trim(), u: unidade(r[C.fil]),
    tipo: r[C.tipo], status: r[C.status], curso: r[C.curso], serie: r[C.serie],
    trilha: r[C.desc], plano: r[C.plano],
    reservas: [], cotas: [], mensalidades: []
  });
  const a = alunos.get(ra);
  const serv = String(r[C.serv] || '');
  const item = { serv, venc: r[C.venc], orig: num(r[C.orig]), bolsa: num(r[C.bolsa]),
    pct: num(r[C.pct]), liq: num(r[C.liq]), idlan: r[C.idlan], parcela: r[C.parcela] };
  if (RESERVA.test(serv)) a.reservas.push(item);
  else if (COTA.test(serv)) a.cotas.push(item);
  else if (MENSALIDADE.test(serv)) a.mensalidades.push(item);
});

const todos = [...alunos.values()];
todos.forEach(a => {
  const b = baixaPorRA.get(a.ra) || { res: 0, cota: 0, dtRes: null, dtCota: null };
  a.pagoRes = b.res; a.pagoCota = b.cota; a.dtRes = b.dtRes; a.dtCota = b.dtCota;
  a.duplaCobranca = a.reservas.length > 0 && a.cotas.length > 0;
  a.aberto = a.duplaCobranca
    ? (b.res > 0 && b.cota === 0 ? 'cota' : b.cota > 0 && b.res === 0 ? 'reserva'
      : b.res > 0 && b.cota > 0 ? 'nenhuma' : 'as duas')
    : null;
  // Em aberto é a soma das parcelas de entrada sem baixa própria — não do
  // grupo inteiro. Uma reserva partida em duas parcelas pode ter só metade paga.
  const semBaixa = x => { const b = baixaPorLan.get(String(x.idlan || '').trim());
    return !(b && b.valor > 0); };
  a.valorAberto = a.duplaCobranca
    ? a.reservas.concat(a.cotas).filter(semBaixa).reduce((s, x) => s + x.liq, 0) : 0;
  // anuidade: soma das mensalidades líquidas mais a entrada efetivamente devida
  const mens = a.mensalidades.reduce((s, x) => s + x.liq, 0);
  a.anuidade = mens > 0 ? mens + (a.reservas[0] ? a.reservas[0].liq : 0) : null;
  a.bolsaPct = a.mensalidades.length ? a.mensalidades[0].pct : (a.cotas[0] ? a.cotas[0].pct : null);
});

const cap = todos.filter(a => a.tipo === 'MATRÍCULA');
const rem = todos.filter(a => a.tipo === 'REMATRÍCULA');
const dupla = cap.filter(a => a.duplaCobranca);

/* ---------- saída ---------- */
const porUnidadeStatus = lista => {
  const m = {};
  lista.forEach(a => { (m[a.u] = m[a.u] || {})[a.status] = ((m[a.u] || {})[a.status] || 0) + 1; });
  return m;
};

const saida = {
  geradoEm: new Date().toISOString().slice(0, 10),
  fontes: { base: path.basename(P_NOVA), baixas: path.basename(P_BAIXA) },
  CODE, ORDER,
  total: todos.length,
  cap: {
    n: cap.length,
    porStatus: cap.reduce((o, a) => (o[a.status] = (o[a.status] || 0) + 1, o), {}),
    porUnidade: porUnidadeStatus(cap)
  },
  rem: {
    n: rem.length,
    porStatus: rem.reduce((o, a) => (o[a.status] = (o[a.status] || 0) + 1, o), {}),
    porUnidade: porUnidadeStatus(rem)
  },
  dupla: {
    n: dupla.length,
    porSituacao: dupla.reduce((o, a) => (o[a.aberto] = (o[a.aberto] || 0) + 1, o), {}),
    valorAberto: dupla.reduce((s, a) => s + a.valorAberto, 0),
    alunos: dupla.map(a => ({
      u: a.u, ra: a.ra, aluno: a.aluno, status: a.status, serie: a.serie, trilha: a.trilha,
      plano: a.plano, aberto: a.aberto, valorAberto: a.valorAberto,
      reserva: a.reservas.map(x => ({ venc: x.venc, liq: x.liq, orig: x.orig })),
      cota: a.cotas.map(x => ({ venc: x.venc, liq: x.liq, orig: x.orig })),
      pagoRes: a.pagoRes, pagoCota: a.pagoCota, dtRes: a.dtRes, dtCota: a.dtCota,
      bolsaPct: a.bolsaPct, anuidade: a.anuidade
    })).sort((x, y) => ORDER.indexOf(x.u) - ORDER.indexOf(y.u) || y.valorAberto - x.valorAberto)
  }
};

const destino = path.join(__dirname, '..', 'dados', 'rastreio-0309.json');
fs.mkdirSync(path.dirname(destino), { recursive: true });
fs.writeFileSync(destino, JSON.stringify(saida, null, 1));

const brl = v => 'R$ ' + Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
console.log('alunos na base:', todos.length, '| captação:', cap.length, '| rematrícula:', rem.length);
console.log('\ndupla cobrança na captação:', dupla.length);
Object.entries(saida.dupla.porSituacao).forEach(([k, v]) => console.log('   ', String(v).padStart(3), 'em aberto:', k));
console.log('valor em aberto:', brl(saida.dupla.valorAberto));
console.log('\ncaptação por status:');
Object.entries(saida.cap.porStatus).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log('   ', String(v).padStart(4), k));
console.log('\nrematrícula por status:');
Object.entries(saida.rem.porStatus).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log('   ', String(v).padStart(4), k));
console.log('\n->', destino);
