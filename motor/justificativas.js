/* A planilha que as unidades preenchem para justificar o que não vai ser
 * corrigido agora.
 *
 *   node motor/justificativas.js [pasta-de-saida]
 *
 * Sai com uma linha por inconsistência aberta no painel do dia: unidade,
 * aluno, RA, o tipo e o que o painel aponta. A unidade escreve a
 * justificativa, o nome de quem escreveu e a data. Na importação seguinte o
 * painel lê de volta e mostra a justificativa ao lado da inconsistência.
 *
 * A cola é RA + tipo, e é só isso que o leitor exige: linha sem RA ou sem
 * justificativa é ignorada, e a ordem das colunas não importa. Deixar o tipo
 * em branco faz a justificativa valer para qualquer inconsistência do aluno.
 *
 * Se já existir uma planilha preenchida, passe em JUSTIFICATIVAS= que o texto
 * de cada linha é carregado para a nova — a unidade não redigita o que já
 * escreveu, e quem foi corrigido no Totvs some da lista sozinho. */

const fs = require('fs');
const path = require('path');
const { pasta, colName } = require('./escrever-xlsx.js');
const { lerJustificativas, ERROS, CODE, ORDER } = require('./cruzar.js');

const B = 'C:/Users/ti/Desktop/Matrículas 2027/';
const P_PAINEL = process.env.PAINEL || path.join(__dirname, '..', 'dados', 'painel.json');
const P_ANTIGA = process.env.JUSTIFICATIVAS || null;
const SAIDA = process.argv[2] || B;

const D = JSON.parse(fs.readFileSync(P_PAINEL, 'utf8'));
const antigas = P_ANTIGA && fs.existsSync(P_ANTIGA)
  ? lerJustificativas(P_ANTIGA).mapa : new Map();

/* Uma linha por inconsistência aberta, na ordem das unidades. */
const linhas = [];
for (const e of D.erros) {
  for (const i of e.itens || []) {
    const ja = i.ra ? (antigas.get(i.ra.toUpperCase() + '|' + e.id) ||
                       antigas.get(i.ra.toUpperCase() + '|*')) : null;
    linhas.push({ u: i.u, unidade: CODE[i.u] || i.u, nome: i.nome, ra: i.ra || '',
      serie: i.serie || '', sit: i.sit || '', chip: e.chip, titulo: e.t,
      texto: ja ? ja.texto : '', quem: ja ? ja.quem : '', data: ja ? ja.data : '' });
  }
}
linhas.sort((a, b) => ORDER.indexOf(a.u) - ORDER.indexOf(b.u) ||
  String(a.nome).localeCompare(String(b.nome), 'pt') || a.chip.localeCompare(b.chip));

/* ---------------- planilha ---------------- */

const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9,
            legenda:10, verde:11, lar:12, larData:13, larMoeda:14, larNum:15 };
const t = (v, s) => ({ v: v == null ? '' : v, t:'s', s: s == null ? S.txt : s });
const n = (v, s) => ({ v: v || 0, t:'n', s: s == null ? S.num : s });

const CAB = ['Unidade','Aluno','RA','Série','Situação','Tipo','O que o painel aponta',
  'Justificativa','Quem escreveu','Data','Prazo'];
const COLS = [18,32,13,8,18,22,46,60,22,12,12];

const L = [{ cells: CAB.map(x => t(x, S.hdr)), h:34 }];
let bloco = null, z = false;
for (const o of linhas) {
  if (bloco !== o.u) { bloco = o.u; z = !z; }
  const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num };
  L.push({ cells: [
    t(o.unidade, st.txt), t(o.nome, st.txt), t(o.ra, st.txt), t(o.serie, st.num),
    t(o.sit, st.txt), t(o.chip, st.num), t(o.titulo, st.txt),
    t(o.texto, o.texto ? st.txt : S.acao),   // a coluna a preencher fica destacada
    t(o.quem, st.txt), t(o.data, st.num), t('', st.num)
  ] });
}

/* ---------------- como preencher ---------------- */

const L2 = [];
[['Como preencher', S.titulo],
 ['Cada linha é uma inconsistência que o painel aponta hoje na sua unidade. Escreva na coluna Justificativa por que ela continua aberta, quem escreveu e a data. As outras colunas já vêm preenchidas: não precisa mexer.', S.legenda],
 ['Filtre pela coluna Unidade para ver só as suas linhas.', S.legenda],
 ['Devolva o arquivo inteiro, com as linhas de todas as unidades. É ele que entra no painel junto com os relatórios do Totvs e da Layers.', S.legenda],
 ['O que a justificativa faz', S.titulo],
 ['A inconsistência continua contada. Ela não some do painel nem do número da unidade — ganha a marca JUSTIFICADA, com o seu texto, o seu nome e há quantos dias foi escrita.', S.legenda],
 ['Justificar não é resolver: serve para quem olha o painel saber por que aquilo ainda está lá, sem precisar perguntar por WhatsApp.', S.legenda],
 ['Se a unidade corrigir no Totvs, o aluno sai da lista sozinho na importação seguinte e a justificativa dorme junto. Se o problema voltar, ela reaparece com a data em que foi escrita.', S.legenda],
 ['Detalhes que importam', S.titulo],
 ['A ligação é feita por RA mais Tipo. Não apague essas duas colunas, e não troque o RA de linha.', S.legenda],
 ['Deixar o Tipo em branco faz a justificativa valer para qualquer inconsistência daquele aluno.', S.legenda],
 ['Linha sem RA ou sem justificativa é ignorada na leitura — pode deixar em branco o que ainda não foi analisado.', S.legenda],
 ['Dá para acrescentar linhas no fim para um aluno que não está na lista. Basta RA, Tipo e Justificativa.', S.legenda],
 ['A ordem das colunas não importa; o painel procura pelo nome do cabeçalho.', S.legenda],
 ['Os tipos que existem', S.titulo]
].forEach(([x, s]) => L2.push({ cells:[t(x, s)], h: s === S.titulo ? 22 : undefined }));
ERROS.forEach(e => L2.push({ cells:[t(e.chip + ' — ' + e.t + '. ' + e.d, S.legenda)] }));

if (require.main === module) {
  const buf = pasta([
    { nome:'Justificativas', cols:COLS, linhas:L, congelar:1,
      filtro: 'A1:' + colName(COLS.length - 1) + L.length },
    { nome:'Como preencher', cols:[130], linhas:L2 }
  ]);
  const destino = path.join(SAIDA, 'Justificativas - preencher.xlsx');
  fs.writeFileSync(destino, buf);

  console.log('->', destino, Math.round(buf.length / 1024) + ' KB');
  console.log(linhas.length + ' inconsistências abertas · painel de ' + D.geradoEm);
  if (antigas.size) console.log('  ' + linhas.filter(o => o.texto).length +
    ' já traziam justificativa da planilha anterior');
  const porU = {};
  linhas.forEach(o => porU[o.u] = (porU[o.u] || 0) + 1);
  ORDER.filter(u => porU[u]).forEach(u =>
    console.log('  ' + (CODE[u] || u).padEnd(20) + String(porU[u]).padStart(4)));
}

module.exports = { linhas };
