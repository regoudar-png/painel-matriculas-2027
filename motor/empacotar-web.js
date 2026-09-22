/* Empacota o motor num único arquivo que roda no navegador.
 *
 *   node motor/empacotar-web.js [saida.js]
 *
 * Junta o leitor de .xlsx do navegador, o cruzamento e a montagem do painel,
 * tirando o que só existe no Node (require, module.exports e os blocos de
 * linha de comando). O resultado expõe MotorPainel.processar(arquivos). */

const fs = require('fs');
const path = require('path');

const M = __dirname;
const ler = n => fs.readFileSync(path.join(M, n), 'utf8');

// Corta do primeiro marco até o fim do arquivo.
function ateAntesDe(s, marco) {
  const i = s.indexOf(marco);
  return i < 0 ? s : s.slice(0, i);
}

/* ---- cruzar.js ---- */
let cruzar = ler('cruzar.js');
cruzar = ateAntesDe(cruzar, "module.exports = { cruzar, CATEGORIAS");
cruzar = cruzar
  .replace("const fs = require('fs');\n", '')
  .replace("const path = require('path');\n", '')
  .replace("const { lerPlanilha } = require('./ler-xlsx.js');",
    "// No navegador a planilha chega já lida; caminho de arquivo não existe.\n" +
    "const lerPlanilha = () => { throw new Error('no navegador a planilha precisa vir já lida'); };");

/* ---- montar-painel.js ---- */
let montar = ler('montar-painel.js');
montar = ateAntesDe(montar, "if (typeof module !== 'undefined' && module.exports) module.exports = { montar");
montar = montar
  .replace("const fs = require('fs');\n", '')
  .replace("const path = require('path');\n", '')
  .replace(/const \{ cruzar, CATEGORIAS[^;]*;\n/, '');

const leitor = ler('ler-xlsx-browser.js');

const saida = `/* Motor do Painel de Matrículas 2027 — versão do navegador.
   Gerado por motor/empacotar-web.js; não editar à mão. */
(function (raiz) {
'use strict';

/* ============ leitor de .xlsx ============ */
${leitor}
const { lerPlanilhaBuf } = raiz.LerXlsx;

/* ============ cruzamento ============ */
${cruzar}
/* ============ montagem do painel ============ */
${montar}
/* ============ entrada ============ */

// arquivos: { totvs:{nome,buf}, marketplace:{...}, financeiro:{...} }
// buf: ArrayBuffer ou Uint8Array. Devolve o JSON que o painel consome.
async function processar(arquivos, aoAndar) {
  const passo = (n, t) => { if (aoAndar) aoAndar(n, t); };
  const ordem = ['totvs', 'marketplace', 'financeiro'];
  const lidos = {};
  let n = 0;
  for (const k of ordem) {
    const a = arquivos[k];
    if (!a || !a.buf) { lidos[k] = null; n++; continue; }
    passo(++n, 'lendo ' + (a.nome || k));
    lidos[k] = await lerPlanilhaBuf(a.buf);
  }
  if (!lidos.totvs) throw new Error('o relatório do Totvs com a base de alunos é obrigatório');
  // A planilha das unidades saiu do processo: o cruzamento roda sem ela.

  passo(4, 'cruzando as bases');
  const r = cruzar(lidos.totvs, null, lidos.marketplace, lidos.financeiro);

  passo(5, 'montando o painel');
  const nomeMkt = arquivos.marketplace && arquivos.marketplace.nome;
  const ate = (String(nomeMkt || '').match(/(\\d{4}-\\d{2}-\\d{2})/) || [])[1] || null;
  const nome = k => (arquivos[k] && arquivos[k].nome) || null;
  return montar(r, ate, { totvs: nome('totvs'), unidades: null,
    marketplace: nome('marketplace'), financeiro: nome('financeiro') });
}

const api = { processar, cruzar, montar, lerPlanilhaBuf, CATEGORIAS, CODE, ORDER };
if (typeof module !== 'undefined' && module.exports) module.exports = api;
raiz.MotorPainel = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
`;

const destino = process.argv[2] || path.join(M, '..', 'painel', 'motor-web.js');
fs.mkdirSync(path.dirname(destino), { recursive: true });
fs.writeFileSync(destino, saida);
console.log('->', destino, Math.round(saida.length / 1024) + ' KB');
