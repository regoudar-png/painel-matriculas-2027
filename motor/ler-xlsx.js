// Leitor de .xlsx sem dependências externas.
// Lê o arquivo direto do caminho: descompacta o zip com o zlib do próprio Node
// (nada de chamar `unzip`, que não existe no servidor) e interpreta o XML.

const fs = require('fs');
const zlib = require('zlib');

/* ---------------- zip ---------------- */

// Lê o diretório central do zip e devolve {nome: Buffer} só das entradas pedidas.
function abrirZip(caminho, querido) {
  const b = fs.readFileSync(caminho);
  // O End of Central Directory fica no fim; o comentário final pode ter até 64KB.
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65558); i--) {
    if (b.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('não parece um .xlsx: fim do zip não encontrado');

  let total = b.readUInt16LE(eocd + 10);
  let tamCD = b.readUInt32LE(eocd + 12);
  let offCD = b.readUInt32LE(eocd + 16);

  // Zip64: quando os campos estouram, os valores reais ficam noutro registro.
  if (total === 0xffff || offCD === 0xffffffff) {
    for (let i = eocd - 20; i >= 0; i--) {
      if (b.readUInt32LE(i) === 0x07064b50) {
        const off64 = Number(b.readBigUInt64LE(i + 8));
        if (b.readUInt32LE(off64) === 0x06064b50) {
          total = Number(b.readBigUInt64LE(off64 + 32));
          tamCD = Number(b.readBigUInt64LE(off64 + 40));
          offCD = Number(b.readBigUInt64LE(off64 + 48));
        }
        break;
      }
    }
  }

  const saida = {};
  let p = offCD;
  for (let n = 0; n < total && p + 46 <= b.length; n++) {
    if (b.readUInt32LE(p) !== 0x02014b50) break;
    const metodo = b.readUInt16LE(p + 10);
    const compTam = b.readUInt32LE(p + 20);
    const nomeTam = b.readUInt16LE(p + 28);
    const extraTam = b.readUInt16LE(p + 30);
    const comentTam = b.readUInt16LE(p + 32);
    const offLocal = b.readUInt32LE(p + 42);
    const nome = b.toString('utf8', p + 46, p + 46 + nomeTam);
    p += 46 + nomeTam + extraTam + comentTam;

    if (querido && !querido(nome)) continue;

    // O cabeçalho local repete nome e extra, com tamanhos próprios.
    const lhNomeTam = b.readUInt16LE(offLocal + 26);
    const lhExtraTam = b.readUInt16LE(offLocal + 28);
    const ini = offLocal + 30 + lhNomeTam + lhExtraTam;
    const bruto = b.subarray(ini, ini + compTam);
    saida[nome] = metodo === 0 ? bruto : zlib.inflateRawSync(bruto);
  }
  return saida;
}

/* ---------------- xml ---------------- */

const desescapar = s => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
  .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
  .replace(/&amp;/g, '&');

// "H" -> 7. Índice baseado em zero, como o array da linha.
function numeroDaColuna(ref) {
  const letras = ref.match(/^[A-Z]+/)[0];
  let n = 0;
  for (const ch of letras) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function lerTextosCompartilhados(xml) {
  if (!xml) return [];
  const out = [];
  const re = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml))) {
    // <si> pode ter vários <t> quando o texto tem formatação por trecho.
    let s = '';
    const tre = /<t[^>]*>([\s\S]*?)<\/t>/g;
    let t;
    while ((t = tre.exec(m[1]))) s += t[1];
    out.push(desescapar(s));
  }
  return out;
}

function lerAba(xml, textos) {
  const linhas = [];
  // O atributo r é opcional no OOXML: quem escreve com SpreadsheetGear (é o
  // caso dos horários exportados do Totvs) omite tanto em <row> quanto em <c>.
  // Sem ele a posição é implícita — a linha seguinte e a próxima coluna.
  const rowRe = /<row\b([^>]*)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let m, linhaAtual = 0;
  while ((m = rowRe.exec(xml))) {
    const attrsL = m[1] || '';
    const rL = (attrsL.match(/\sr="(\d+)"/) || [])[1];
    const numero = rL ? +rL : linhaAtual + 1;
    linhaAtual = numero;
    const corpo = m[2] || '';
    const arr = [];
    const cRe = /<c\b([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let c, colAtual = -1;
    while ((c = cRe.exec(corpo))) {
      const attrs = c[1], dentro = c[2] || '';
      const ref = (attrs.match(/\sr="([A-Z]+\d+)"/) || [])[1];
      const col = ref ? numeroDaColuna(ref) : colAtual + 1;
      colAtual = col;
      const tipo = (attrs.match(/\st="([^"]+)"/) || [])[1];
      let v = null;
      if (tipo === 'inlineStr') {
        let s = '';
        const tre = /<t[^>]*>([\s\S]*?)<\/t>/g;
        let t;
        while ((t = tre.exec(dentro))) s += t[1];
        v = desescapar(s);
      } else {
        const vm = dentro.match(/<v>([\s\S]*?)<\/v>/);
        if (vm) v = tipo === 's' ? textos[+vm[1]] : desescapar(vm[1]);
      }
      arr[col] = v;
    }
    linhas[numero - 1] = arr;
  }
  for (let i = 0; i < linhas.length; i++) if (!linhas[i]) linhas[i] = [];
  return linhas;
}

/* ---------------- api ---------------- */

/**
 * Abre um .xlsx e devolve { abas: [{nome, linhas}] }.
 * `linhas` é um array de arrays; célula ausente vira `undefined`.
 */
function lerPlanilha(caminho, opcoes = {}) {
  const querido = n =>
    n === 'xl/workbook.xml' || n === 'xl/_rels/workbook.xml.rels' ||
    n === 'xl/sharedStrings.xml' || n.startsWith('xl/worksheets/');
  const arq = abrirZip(caminho, querido);

  const wb = arq['xl/workbook.xml'];
  if (!wb) throw new Error('xl/workbook.xml ausente em ' + caminho);
  const rels = (arq['xl/_rels/workbook.xml.rels'] || Buffer.from('')).toString('utf8');

  const alvoPorId = {};
  let m;
  const relRe = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  while ((m = relRe.exec(rels))) alvoPorId[m[1]] = m[2];

  const textos = lerTextosCompartilhados(
    arq['xl/sharedStrings.xml'] && arq['xl/sharedStrings.xml'].toString('utf8'));

  const abas = [];
  const sheetRe = /<sheet[^>]*\/>/g;
  const wbXml = wb.toString('utf8');
  while ((m = sheetRe.exec(wbXml))) {
    const tag = m[0];
    const nome = desescapar((tag.match(/name="([^"]*)"/) || [])[1] || '');
    const rid = (tag.match(/r:id="([^"]*)"/) || [])[1];
    const alvo = alvoPorId[rid];
    if (!alvo) continue;
    const chave = 'xl/' + alvo.replace(/^\//, '').replace(/^xl\//, '');
    const dados = arq[chave];
    if (!dados) continue;
    if (opcoes.apenas && !opcoes.apenas.includes(nome)) { abas.push({ nome, linhas: null }); continue; }
    abas.push({ nome, linhas: lerAba(dados.toString('utf8'), textos) });
  }
  return { abas };
}

module.exports = { lerPlanilha, numeroDaColuna, desescapar };
