/* Lista dos alunos com a entrada cobrada duas vezes.
 *
 *   node motor/lista-entrada-dobrada.js [pasta-de-saida]
 *
 * Sai direto de dados/painel.json — o mesmo resultado que alimenta o painel
 * publicado, para que a planilha e o link nunca digam números diferentes.
 * Três abas: uma linha por aluno com o que a unidade precisa resolver, as
 * cobranças cruas que geraram o apontamento, e o resumo por unidade. */

const fs = require('fs');
const path = require('path');
const { pasta, colName } = require('./escrever-xlsx.js');

const RAIZ = path.join(__dirname, '..');
const P_DADOS = path.join(RAIZ, 'dados', 'painel.json');
const SAIDA = process.argv[2] || 'C:/Users/ti/Desktop/Matrículas 2027';

const D = JSON.parse(fs.readFileSync(P_DADOS, 'utf8'));
const CODE = D.CODE, ORDER = D.ORDER;
const lista = D.erros.find(e => e.id === 'cota_duplicada');
if (!lista) { console.error('a lista de entrada em dobro não está no painel'); process.exit(1); }

/* ---------------- nomes ---------------- */
// O painel guarda o nome em caixa baixa para casar as bases; para ler numa
// planilha ele precisa voltar à caixa de nome próprio.
const PARTICULA = new Set(['de', 'da', 'do', 'das', 'dos', 'e', 'di', 'du', 'del', 'la', 'van', 'von', 'y']);
// Parte da base vem em caixa alta e parte em caixa baixa: baixo tudo antes de
// levantar a inicial, senão sobram nomes gritando no meio da lista.
const nomeProprio = s => String(s || '').trim().toLowerCase().split(/\s+/)
  .map((p, i) => (i > 0 && PARTICULA.has(p)) ? p : p.charAt(0).toUpperCase() + p.slice(1))
  .join(' ');

/* ---------------- células ---------------- */
const S = { hdr: 1, txt: 2, data: 3, moeda: 4, pct: 5, num: 6, acao: 7, marca: 8, titulo: 9,
            legenda: 10, verde: 11, lar: 12, larData: 13, larMoeda: 14, larNum: 15 };
const t = (v, s) => ({ v, t: 's', s: s == null ? S.txt : s });
const n = (v, s) => ({ v, t: 'n', s: s == null ? S.num : s });
const m = (v, s) => ({ v, t: 'm', s: s == null ? S.moeda : s });
const brl = v => 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* ---------------- os 104 ---------------- */
const itens = lista.itens.slice().sort((a, b) =>
  ORDER.indexOf(a.u) - ORDER.indexOf(b.u) || a.nome.localeCompare(b.nome, 'pt'));

const daEntrada = e => {
  const s = String(e.servico || '');
  if (/Reserva de vaga/i.test(s)) return 'Reserva de vaga';
  if (/1ª Cota/i.test(s)) return '1ª cota do contrato';
  return s.replace(/^[A-Z0-9]+ - /, '');
};

/* ================= aba 1: os alunos ================= */
const CAB1 = ['#', 'Unidade', 'Aluno', 'Série', 'RA', 'Situação', 'Cobranças', 'Já baixado',
  'Em aberto (a resolver)', 'Anuidade', 'Plano do ano', 'O que fazer', 'Resolvido?', 'Observação da unidade'];
const L1 = [{ cells: CAB1.map(x => t(x, S.hdr)), h: 30 }];

itens.forEach((i, k) => {
  const baixadas = (i.entradas || []).filter(e => e.baixado);
  const abertas = (i.entradas || []).filter(e => !e.baixado);
  const pago = baixadas.reduce((s, e) => s + (e.valor || 0), 0);
  const zebra = k % 2 ? 1 : 0;   // a faixa clara marca a linha ímpar
  const st = { txt: zebra ? S.lar : S.txt, num: zebra ? S.larNum : S.num,
               moeda: zebra ? S.larMoeda : S.moeda };
  // O que fazer sai do próprio caso: com a entrada já paga, o que sobra é
  // cancelar a segunda; sem nenhuma paga, é escolher qual cobrar.
  const acao = !abertas.length ? 'Conferir: as duas entradas foram baixadas'
    : !baixadas.length ? 'Nenhuma foi paga: cancelar uma das duas'
    : 'Cancelar ' + abertas.map(daEntrada).join(' e ') + ' (ou virar desconto)';
  L1.push({ cells: [
    n(k + 1, st.num),
    t(CODE[i.u] || i.u, st.txt),
    t(nomeProprio(i.nome), st.txt),
    t(i.serie || '', st.num),
    t(i.ra || '', st.txt),
    t(i.sit || '', st.txt),
    n(i.nEntradas || (i.entradas || []).length, st.num),
    m(pago, st.moeda),
    m(i.duplicado || 0, st.moeda),
    i.anual != null ? m(i.anual, st.moeda) : t('—', st.num),
    t(i.temMensalidade ? 'Sim' : 'Não', st.num),
    t(acao, st.txt),
    t('', S.acao),
    t('', st.txt)
  ] });
});

const totalAberto = itens.reduce((s, i) => s + (i.duplicado || 0), 0);
L1.push({ cells: [t(''), t('Total', S.verde), t(itens.length + ' alunos', S.verde), t(''), t(''), t(''),
  t(''), t(''), t(brl(totalAberto), S.verde)] });

/* ================= aba 2: as cobranças ================= */
const CAB2 = ['Unidade', 'Aluno', 'RA', 'Cobrança', 'Serviço no Totvs', 'Vencimento', 'Valor', 'Baixado'];
const L2 = [{ cells: CAB2.map(x => t(x, S.hdr)), h: 24 }];
let linha = 0;
itens.forEach(i => {
  (i.entradas || []).forEach(e => {
    const zebra = linha++ % 2 ? 1 : 0;
    const st = { txt: zebra ? S.lar : S.txt, num: zebra ? S.larNum : S.num,
                 moeda: zebra ? S.larMoeda : S.moeda };
    L2.push({ cells: [
      t(CODE[i.u] || i.u, st.txt),
      t(nomeProprio(i.nome), st.txt),
      t(i.ra || '', st.txt),
      t(daEntrada(e), st.txt),
      t(e.servico || '', st.txt),
      t(e.venc || '', st.num),
      m(e.valor || 0, st.moeda),
      t(e.baixado ? 'Sim' : 'Não', st.num)
    ] });
  });
});

/* ================= aba 3: por unidade ================= */
const CAB3 = ['Unidade', 'Alunos com entrada em dobro', 'Matrículas da unidade', '% da unidade', 'Valor em aberto'];
const L3 = [{ cells: CAB3.map(x => t(x, S.hdr)), h: 24 }];
const porUni = ORDER.map(c => {
  const it = itens.filter(i => i.u === c);
  const p = D.cap.placar.find(x => x.code === c);
  return { c, n: it.length, mat: p ? p.mat : 0,
           v: it.reduce((s, i) => s + (i.duplicado || 0), 0) };
}).filter(x => x.n).sort((a, b) => b.n - a.n);
porUni.forEach((u, k) => {
  const zebra = k % 2 ? 1 : 0;
  const st = { txt: zebra ? S.lar : S.txt, num: zebra ? S.larNum : S.num,
               moeda: zebra ? S.larMoeda : S.moeda };
  L3.push({ cells: [t(CODE[u.c] || u.c, st.txt), n(u.n, st.num), n(u.mat, st.num),
    { v: u.mat ? u.n / u.mat : 0, t: 'p', s: zebra ? S.larNum : S.pct }, m(u.v, st.moeda)] });
});
// O estilo do total é negrito verde e não carrega formato de número: o
// percentual da rede vai como texto, senão sai 0,389513109 na cara da linha.
const pcRede = D.cap.matriculados ? (itens.length / D.cap.matriculados * 100) : 0;
L3.push({ cells: [t('Rede', S.verde), n(itens.length, S.verde), n(D.cap.matriculados, S.verde),
  t(pcRede.toFixed(2).replace('.', ',') + '%', S.verde),
  t(brl(totalAberto), S.verde)] });
L3.push({ cells: [] });
L3.push({ cells: [t('Apurado em ' + D.geradoEm.split('-').reverse().join('/') +
  ' · base do Totvs ' + (D.fontes.totvs || '?') + ' · ficha financeira ' +
  (D.fontes.financeiro || '?'), S.legenda)] });
L3.push({ cells: [t('Entrada em dobro é mais de um lançamento de entrada no mesmo RA: a reserva de vaga ' +
  'do bolsão e a 1ª cota gerada pelo contrato, ou a mesma cobrança repetida. A família paga uma; a outra ' +
  'fica em aberto e precisa ser cancelada ou virar desconto.', S.legenda)] });

/* ---------------- grava ---------------- */
const buf = pasta([
  { nome: 'Entrada em dobro', cols: [5, 17, 32, 7, 14, 13, 11, 13, 17, 13, 11, 44, 12, 34], linhas: L1, congelar: 1,
    filtro: 'A1:' + colName(CAB1.length - 1) + L1.length,
    validacao: { ref: 'M2:M' + (L1.length - 1), opcoes: ['Sim', 'Não', 'Em análise'],
      dica: 'A unidade marca aqui quando a cobrança em dobro for resolvida.' },
    formatacao: { ref: 'M2:M' + (L1.length - 1), igual: 'Sim' } },
  { nome: 'Cobranças', cols: [17, 32, 14, 20, 42, 13, 12, 10], linhas: L2, congelar: 1,
    filtro: 'A1:' + colName(CAB2.length - 1) + L2.length },
  { nome: 'Por unidade', cols: [22, 26, 22, 12, 17], linhas: L3, congelar: 1 }
]);
const destino = path.join(SAIDA, 'Entrada cobrada duas vezes - captação 2027.xlsx');
fs.writeFileSync(destino, buf);
console.log('->', destino, Math.round(buf.length / 1024) + ' KB');
console.log(itens.length + ' alunos ·', L2.length - 1, 'cobranças ·', brl(totalAberto), 'em aberto');
porUni.forEach(u => console.log('  ' + (CODE[u.c] || u.c).padEnd(20), String(u.n).padStart(3), brl(u.v)));
