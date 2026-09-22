// Lista as cotas pagas pelo Portal do aluno, a partir do snapshot já cruzado.
//
//   node motor/lista-portal.js [pasta-de-saida]

const fs = require('fs');
const path = require('path');
const { pasta, colName } = require('./escrever-xlsx.js');

const SAIDA = process.argv[2] || '.';
const D = require('../dados/painel.json');
const CODE = D.CODE, ORDER = D.ORDER;

const itens = [];
D.cap.listas.forEach(l => l.itens.forEach(i => itens.push({ ...i, categoria: l.t })));

// Pago pelo Portal = tem baixa no financeiro do Totvs e não veio do Marketplace.
// Inclui os parciais, marcados como tal — a unidade precisa vê-los.
const portal = itens
  .filter(i => (i.pagoP || i.parcial) && !i.pagoM)
  .sort((a, b) => ORDER.indexOf(a.u) - ORDER.indexOf(b.u) ||
                  String(a.nome).localeCompare(String(b.nome), 'pt'));

const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9, legenda:10, verde:11 };

const COLS = [
  ['Unidade', 20], ['Sigla', 7], ['Aluno', 34], ['RA', 14], ['Série', 12],
  ['Situação no Totvs', 24], ['Cota esperada', 13], ['Valor pago', 12],
  ['Falta', 10], ['Status da cota', 15], ['Data da baixa', 13],
  ['Situação apurada', 34]
];

function linha(i) {
  const c = v => ({ v: v == null ? '' : String(v), t: 's', s: S.txt });
  const m = v => v == null ? { v:'', t:'s', s:S.moeda } : { v, t:'m', s:S.moeda };
  const completa = !i.parcial;
  return { cells: [
    c(CODE[i.u]), c(i.u), c(i.nome), c(i.ra), c(i.serie),
    c(i.sit || 'fora do Totvs'),
    m(i.cotaEsperada), m(i.cotaEsperada != null && i.falta != null ? i.cotaEsperada - i.falta : i.cotaEsperada),
    m(i.parcial ? i.falta : null),
    { v: completa ? 'COMPLETA' : 'PARCIAL', t:'s', s: completa ? S.verde : S.marca },
    i.finDt ? { v: i.finDt.split('/').reverse().join('-'), t:'d', s:S.data } : c(''),
    c(i.categoria)
  ], h: 20 };
}

function resumo() {
  const L = [];
  const t = (v, s) => ({ v, t:'s', s: s || S.txt });
  L.push({ cells: [{ v:'Cotas pagas pelo Portal do aluno', t:'s', s:S.titulo }], h:24 });
  L.push({ cells: [{ v:'Matrículas 2027 · fonte: financeiro do Totvs · '+(D.fontes.financeiro||''), t:'s', s:S.legenda }] });
  L.push({ cells: [{ v:'Cota de entrada: R$ 300 nas unidades e R$ 600 em Américas. '+
    'A mensalidade não entra nesta conta.', t:'s', s:S.legenda }], h:28 });
  L.push({ cells: [] });
  L.push({ cells: [t('Unidade', S.hdr), t('Cota completa', S.hdr), t('Cota parcial', S.hdr),
                   t('Total', S.hdr), t('Valor recebido', S.hdr)], h:26 });
  const porUni = {};
  portal.forEach(i => {
    const u = porUni[i.u] = porUni[i.u] || { comp:0, parc:0, valor:0 };
    if (i.parcial) u.parc++; else u.comp++;
    u.valor += (i.cotaEsperada||0) - (i.falta||0);
  });
  ORDER.filter(u => porUni[u]).forEach(u => {
    const v = porUni[u];
    L.push({ cells: [t(CODE[u]),
      { v:v.comp, t:'n', s:S.num }, { v:v.parc, t:'n', s:S.num },
      { v:v.comp+v.parc, t:'n', s:S.num }, { v:v.valor, t:'m', s:S.moeda }] });
  });
  const tot = Object.values(porUni).reduce((a,v)=>({comp:a.comp+v.comp, parc:a.parc+v.parc, valor:a.valor+v.valor}),
    {comp:0,parc:0,valor:0});
  L.push({ cells: [t('Rede', S.hdr),
    { v:tot.comp, t:'n', s:S.hdr }, { v:tot.parc, t:'n', s:S.hdr },
    { v:tot.comp+tot.parc, t:'n', s:S.hdr }, { v:tot.valor, t:'m', s:S.moeda }] });
  return { nome:'Resumo', cols:[24,14,14,10,16], linhas:L, congelar:0 };
}

const linhas = [{ cells: COLS.map(h => ({ v:h[0], t:'s', s:S.hdr })), h:28 }]
  .concat(portal.map(linha));
const ultima = colName(COLS.length-1) + linhas.length;

const abas = [
  resumo(),
  { nome:'Cotas pelo Portal', cols:COLS.map(h=>h[1]), linhas, congelar:1, filtro:'A1:'+ultima }
];

const arq = path.join(SAIDA, 'Cotas pagas pelo Portal - 2027.xlsx');
fs.writeFileSync(arq, pasta(abas));

const comp = portal.filter(i=>!i.parcial).length;
console.log('cotas pelo Portal:', portal.length, '·', comp, 'completas ·', portal.length-comp, 'parciais');
const porUni = {};
portal.forEach(i => porUni[i.u] = (porUni[i.u]||0)+1);
console.log('por unidade:', ORDER.filter(u=>porUni[u]).map(u=>u+':'+porUni[u]).join('  '));
console.log('->', arq, Math.round(fs.statSync(arq).size/1024)+' KB');
