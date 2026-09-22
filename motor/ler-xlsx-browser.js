/* Leitor de .xlsx para o navegador.
 *
 * Mesma lógica do ler-xlsx.js do Node, com duas trocas: o buffer vira
 * Uint8Array/DataView e o inflate sai do zlib para o DecompressionStream do
 * próprio navegador — por isso tudo aqui é assíncrono.
 *
 * Roda também no Node (DecompressionStream existe desde a 18), o que permite
 * conferir o resultado contra o leitor original antes de publicar. */

(function (raiz) {
  'use strict';

  const TD = new TextDecoder('utf-8');
  const texto = (u8, ini, fim) => TD.decode(u8.subarray(ini, fim));

  async function inflarBruto(u8) {
    const ds = new DecompressionStream('deflate-raw');
    const fluxo = new Blob([u8]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(fluxo).arrayBuffer());
  }

  /* ---------------- zip ---------------- */

  async function abrirZip(u8, querido) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const u32 = p => dv.getUint32(p, true);
    const u16 = p => dv.getUint16(p, true);
    const u64 = p => Number(dv.getBigUint64(p, true));

    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65558); i--) {
      if (u32(i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('não parece um .xlsx: fim do zip não encontrado');

    let total = u16(eocd + 10);
    let offCD = u32(eocd + 16);

    // Zip64: quando os campos estouram, os valores reais ficam noutro registro.
    if (total === 0xffff || offCD === 0xffffffff) {
      for (let i = eocd - 20; i >= 0; i--) {
        if (u32(i) === 0x07064b50) {
          const off64 = u64(i + 8);
          if (u32(off64) === 0x06064b50) {
            total = u64(off64 + 32);
            offCD = u64(off64 + 48);
          }
          break;
        }
      }
    }

    const saida = {};
    let p = offCD;
    for (let n = 0; n < total && p + 46 <= u8.length; n++) {
      if (u32(p) !== 0x02014b50) break;
      const metodo = u16(p + 10);
      const compTam = u32(p + 20);
      const nomeTam = u16(p + 28);
      const extraTam = u16(p + 30);
      const comentTam = u16(p + 32);
      const offLocal = u32(p + 42);
      const nome = texto(u8, p + 46, p + 46 + nomeTam);
      p += 46 + nomeTam + extraTam + comentTam;

      if (querido && !querido(nome)) continue;

      const lhNomeTam = u16(offLocal + 26);
      const lhExtraTam = u16(offLocal + 28);
      const ini = offLocal + 30 + lhNomeTam + lhExtraTam;
      const bruto = u8.subarray(ini, ini + compTam);
      saida[nome] = metodo === 0 ? bruto : await inflarBruto(bruto);
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
    // caso dos horários exportados do Totvs) omite tanto em <row> quanto em
    // <c>. Sem ele a posição é implícita — a linha seguinte e a próxima coluna.
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

  // Recebe ArrayBuffer ou Uint8Array; devolve { abas: [{nome, linhas}] }.
  async function lerPlanilhaBuf(entrada, opcoes) {
    opcoes = opcoes || {};
    const u8 = entrada instanceof Uint8Array ? entrada : new Uint8Array(entrada);
    const querido = n =>
      n === 'xl/workbook.xml' || n === 'xl/_rels/workbook.xml.rels' ||
      n === 'xl/sharedStrings.xml' || n.startsWith('xl/worksheets/');
    const arq = await abrirZip(u8, querido);

    const wb = arq['xl/workbook.xml'];
    if (!wb) throw new Error('xl/workbook.xml ausente: o arquivo não é um .xlsx válido');
    const rels = arq['xl/_rels/workbook.xml.rels'] ? TD.decode(arq['xl/_rels/workbook.xml.rels']) : '';

    const alvoPorId = {};
    let m;
    const relRe = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
    while ((m = relRe.exec(rels))) alvoPorId[m[1]] = m[2];

    const textos = lerTextosCompartilhados(
      arq['xl/sharedStrings.xml'] ? TD.decode(arq['xl/sharedStrings.xml']) : null);

    const abas = [];
    const wbXml = TD.decode(wb);
    const sheetRe = /<sheet[^>]*\/>/g;
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
      abas.push({ nome, linhas: lerAba(TD.decode(dados), textos) });
    }
    return { abas };
  }

  const api = { lerPlanilhaBuf, numeroDaColuna, desescapar };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  raiz.LerXlsx = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
