/* A entrada da captação 2027, lançamento por lançamento.
 *
 *   node motor/entrada-2027.js [pasta-de-saida]
 *
 * O desenho correto é cota 1 de R$ 300 e 12 mensalidades. A ficha financeira
 * mostra que a 1ª Cota de Mensalidade NÃO é uma das 12: ela é a PARCELA 0 do
 * contrato, e as mensalidades são as parcelas 1 a 12, competências 01 a 12 de
 * 2027. QTD_PARC_CONTRATO = 13 confirma: 1 cota + 12 mensalidades. Cancelar a
 * cota, portanto, deixa as 12 mensalidades de pé.
 *
 * Uma entrada pode vir partida. QTD_PARC_CONTRATO − 12 diz em quantas parcelas:
 * 13 → 1, 14 → 2, 18 → 6. Parcelas da mesma entrada não são cobrança em dobro —
 * foi aí que a lista anterior errou, marcando para cancelar a 2ª parcela de
 * quem tinha a entrada partida.
 *
 * Também não são cobrança em dobro:
 *   · o lançamento cancelado (STATUS_BOLETO = Cancelado);
 *   · a cota zerada com bolsa e baixada (Baixado com VALOR_BAIXADO = 0);
 *   · a linha sem lançamento financeiro, que vem do export com VALOR_ORIGINAL
 *     vazio, REF_FINANCEIRO 4837 e VALOR_BAIXADO 4838 em todas — número de
 *     controle do relatório, não dinheiro.
 *
 * O que sobra são dois problemas diferentes:
 *   1. duas entradas vivas: a reserva de vaga do bolsão e a 1ª cota do contrato;
 *   2. a 1ª cota emitida pelo valor cheio da tabela, com os R$ 300 do bolsão
 *      baixados por dentro e o saldo em aberto. */

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');
const { pasta, colName } = require('./escrever-xlsx.js');

const B = 'C:/Users/ti/Desktop/Matrículas 2027/';
const P_FICHA = B + '09 de setembro/Ficha financeira.XLSX';
const SAIDA = process.argv[2] || B;

const ENTRADA_ESPERADA = 300;
const MENSALIDADES_ESPERADAS = 12;

/* ---------------- unidades ---------------- */
const CODE = {
  BG:'Bangu', CG:'Campo Grande', CX:'Duque de Caxias', MD:'Madureira', NI:'Nova Iguaçu',
  RM:'Rocha Miranda', RT:'Retiro dos Artistas', SJ:'São João de Meriti', TQ:'Taquara',
  TJ:'Tijuca', AM:'Américas'
};
const ORDER = ['CG','AM','BG','RM','CX','MD','TQ','TJ','SJ','NI','RT'];
const DE_FILIAL = {
  'TAQUARA':'TQ', 'DUQUE DE CAXIAS':'CX', 'BANGU':'BG', 'ROCHA MIRANDA':'RM', 'NOVA IGUACU':'NI',
  'SAO JOAO DE MERITI':'SJ', 'CAMPO GRANDE':'CG', 'RETIRO DOS ARTISTAS':'RT', 'TIJUCA':'TJ',
  'MADUREIRA':'MD', 'RECREIO':'AM'
};
const semAcento = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const unidade = f => DE_FILIAL[semAcento(f).toUpperCase()
  .replace(/COLEGIO E CURSO MATRIZ EDUCACAO ?-? ?/, '').trim()] || '??';

/* ---------------- leitura ---------------- */
const num = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; };
const ehData = v => /^\d{2}\/\d{2}\/\d{4}$/.test(String(v || '').trim());
const TESTE = /\bteste?s?\b/i;
const PARTICULA = new Set(['de','da','do','das','dos','e','di','du','del','la','van','von','y']);
const nomeProprio = s => String(s || '').trim().toLowerCase().split(/\s+/)
  .map((p, i) => (i > 0 && PARTICULA.has(p)) ? p : p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
const brl = v => 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits:2, maximumFractionDigits:2 });
const perto = (a, b) => Math.abs(a - b) < 0.5;

const RESERVA = /Reserva de vaga/i, COTA = /1ª Cota/i, MENS = /(^|- )Mensalidade$/i;

const L = lerPlanilha(P_FICHA).abas.filter(a => a.linhas)
  .sort((a, b) => b.linhas.length - a.linhas.length)[0].linhas;
const cap = L.slice(1).filter(r => r && r[3] && r[19] === 'MATRÍCULA' && !TESTE.test(String(r[4] || '')));

const fantasma = r => num(r[42]) < 0.005 && String(r[47] || '') === 'SEM BOLETO';

// A mesma parcela vem duas vezes quando o boleto tem dois status de remessa.
const vistos = new Set();
const linhas = cap.filter(r => {
  const ref = String(r[35] || '').trim();
  if (!ref || fantasma(r)) return true;
  const k = r[3] + '|' + ref;
  if (vistos.has(k)) return false;
  vistos.add(k); return true;
});

const alunos = new Map();
linhas.forEach(r => {
  if (!alunos.has(r[3])) alunos.set(r[3], {
    ra: r[3], aluno: r[4], u: unidade(r[2]), sit: r[21] || '', serie: r[17] || '',
    plano: r[25] || '', contrato: r[26] || '', qtdParc: num(r[27]),
    lanc: [], mens: 0, mensCanc: 0 });
  const a = alunos.get(r[3]);
  const s = String(r[34] || '');
  const cancelado = String(r[47] || '') === 'Cancelado';
  if (MENS.test(s)) { cancelado ? a.mensCanc++ : a.mens++; return; }
  if (!RESERVA.test(s) && !COTA.test(s)) return;
  const o = {
    grupo: RESERVA.test(s) ? 'Reserva de vaga' : '1ª cota do contrato',
    servico: s, parcela: num(r[28]), venc: String(r[41] || '').trim(),
    orig: num(r[42]), bolsa: num(r[43]), baixado: num(r[44]),
    dt: ehData(r[45]) ? String(r[45]).trim() : null,
    st: String(r[47] || ''), bolsas: String(r[48] || '').trim(),
    ref: String(r[35] || '').trim(), fantasma: fantasma(r)
  };
  o.liq = o.orig - o.bolsa;
  a.lanc.push(o);
});

/* ---------------- situação de cada lançamento ---------------- */
const SIT = {
  fantasma: 'Sem lançamento financeiro',
  cancelada:'Já cancelada',
  quitada:  'Quitada',
  parcial:  'Baixa parcial',
  bolsa:    'Zerada com bolsa e baixada',
  aberta:   'Em aberto'
};
const situacaoDe = o => o.fantasma ? SIT.fantasma
  : o.st === 'Cancelado' ? SIT.cancelada
  : o.baixado > 0 ? (o.baixado >= o.liq - 0.5 ? SIT.quitada : SIT.parcial)
  : o.st === 'Baixado' ? SIT.bolsa
  : SIT.aberta;

alunos.forEach(a => {
  a.lanc.forEach(o => { o.sitCob = situacaoDe(o); });

  // Um grupo é uma entrada: a reserva de vaga é uma, a 1ª cota é outra. As
  // parcelas de uma entrada partida ficam no mesmo grupo, e por isso não se
  // confundem com cobrança em dobro.
  const vivos = a.lanc.filter(o => o.sitCob !== SIT.cancelada && o.sitCob !== SIT.fantasma);
  const grupos = [];
  ['Reserva de vaga', '1ª cota do contrato'].forEach(g => {
    const ls = vivos.filter(o => o.grupo === g);
    if (!ls.length) return;
    grupos.push({ nome: g, ls,
      cobrado: ls.reduce((s, o) => s + o.liq, 0),
      pago: ls.reduce((s, o) => s + o.baixado, 0),
      resolvidoPorBolsa: ls.every(o => o.sitCob === SIT.bolsa) });
  });
  grupos.forEach(g => { g.emAberto = g.resolvidoPorBolsa ? 0 : g.cobrado - g.pago; });
  a.grupos = grupos;

  // Quantas parcelas a entrada deveria ter: o contrato diz, tirando as 12
  // mensalidades. Sem isso a entrada partida em 6 parece seis cobranças.
  a.parcelasEsperadas = a.qtdParc > MENSALIDADES_ESPERADAS
    ? a.qtdParc - MENSALIDADES_ESPERADAS : 1;

  const pagos = grupos.filter(g => g.pago > 0.005 || g.resolvidoPorBolsa);
  const principal = pagos[0] || grupos[0] || null;
  a.principal = principal;
  a.duplicado = grupos.length > 1;

  a.lanc.forEach(o => {
    if (o.sitCob === SIT.fantasma) {
      o.acao = 'Linha sem lançamento financeiro: não é cobrança, não há o que cancelar'; return; }
    if (o.sitCob === SIT.cancelada) { o.acao = 'Nada a fazer: já cancelada'; return; }
    if (o.sitCob === SIT.bolsa) { o.acao = 'Nada a fazer: a unidade zerou com bolsa e baixou'; return; }
    const g = grupos.find(x => x.nome === o.grupo);

    // Duas entradas vivas: a que não recebeu dinheiro é a que sai.
    if (a.duplicado && principal && g !== principal) {
      if (o.sitCob === SIT.aberta) { o.acao = 'CANCELAR: entrada em dobro, a outra já foi paga'; o.cancelar = o.liq; return; }
      o.acao = 'Entrada em dobro com as duas pagas: devolver ou creditar'; o.devolver = o.baixado; return;
    }
    if (a.duplicado && (!principal || g === principal) && o.sitCob === SIT.aberta && grupos.length > 1) {
      o.acao = 'Duas entradas abertas e nenhuma paga: manter uma e cancelar a outra'; return; }

    if (o.sitCob === SIT.quitada) { o.acao = 'Nada a fazer'; return; }
    if (o.sitCob === SIT.parcial) {
      // Os R$ 300 do bolsão entraram por dentro de uma cota emitida cheia.
      const resto = o.liq - o.baixado;
      o.acao = perto(o.baixado, ENTRADA_ESPERADA)
        ? 'CANCELAR o saldo de ' + brl(resto) + ': a entrada de R$ 300 já foi paga por dentro desta cota'
        : 'Baixa parcial de ' + brl(o.baixado) + ': conferir o saldo de ' + brl(resto);
      if (perto(o.baixado, ENTRADA_ESPERADA)) o.cancelar = resto;
      return;
    }
    // Em aberto, entrada única
    if (g && g.pago > 0.005) { o.acao = 'Parcela ' + o.parcela + ' de ' + a.parcelasEsperadas +
      ' da mesma entrada: cobrar, não cancelar'; return; }
    o.acao = 'Entrada ainda não paga: cobrar ou conferir pagamento na Layers';
  });

  /* desvio do padrão: cota 1 de R$ 300 + 12 mensalidades */
  const fora = [];
  // A dobra só é desvio enquanto não foi tratada. Reserva paga com a cota
  // zerada por bolsa e baixada é a dobra já resolvida pela unidade — 49 casos
  // que, listados aqui, virariam trabalho refeito.
  const dobraViva = a.duplicado && grupos.some(g => g !== principal && !g.resolvidoPorBolsa);
  if (dobraViva) fora.push('duas entradas vivas: reserva de vaga e 1ª cota');
  if (!grupos.length) fora.push('sem nenhuma cobrança de entrada');
  else {
    const g = principal || grupos[0];
    if (g.ls.length !== a.parcelasEsperadas)
      fora.push(g.ls.length + ' parcelas de entrada, o contrato prevê ' + a.parcelasEsperadas);
    // O que interessa aqui é o que a família pagou de entrada. A cota emitida
    // pelo valor cheio já é tratada como saldo a cancelar na aba de ajuste;
    // repetir isso aqui encheria a lista com quase toda a rede.
    if (!g.resolvidoPorBolsa) {
      if (g.pago < 0.005) fora.push('entrada ainda não paga');
      else if (!perto(g.pago, ENTRADA_ESPERADA))
        fora.push('entrada paga de ' + brl(g.pago) + ', não de R$ 300');
    }
  }
  const esperaPlano = /^(Matriculado|Pré-Matriculado)$/.test(a.sit);
  if (a.mens === 0) { if (esperaPlano) fora.push('sem mensalidade lançada'); }
  else if (a.mens > MENSALIDADES_ESPERADAS) fora.push(a.mens + ' mensalidades: plano do ano em dobro');
  else if (a.mens < MENSALIDADES_ESPERADAS) fora.push('só ' + a.mens + ' mensalidades');
  a.fora = fora;
});

/* ---------------- células ---------------- */
const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9,
            legenda:10, verde:11, lar:12, larData:13, larMoeda:14, larNum:15 };
const t = (v, s) => ({ v, t:'s', s: s == null ? S.txt : s });
const n = (v, s) => ({ v, t:'n', s: s == null ? S.num : s });
const m = (v, s) => ({ v, t:'m', s: s == null ? S.moeda : s });

const ordenados = [...alunos.values()].sort((a, b) =>
  ORDER.indexOf(a.u) - ORDER.indexOf(b.u) || String(a.aluno).localeCompare(String(b.aluno), 'pt'));

const CAB = ['Unidade','Aluno','RA','Série','Situação do aluno','Entrada','Serviço no Totvs','Parcela',
  'Parcelas previstas','Vencimento','Valor original','Bolsa / deduções','Valor líquido','Valor baixado',
  'Saldo em aberto','Data da baixa','Status do boleto','Bolsas cadastradas','Ref. financeiro',
  'Situação da cobrança','O que fazer','A cancelar','Mensalidades','Resolvido?','Observação da unidade'];
const COLS = [17,30,13,9,24,20,34,8,12,12,13,14,13,13,14,12,13,34,12,24,58,13,12,12,30];

function abaCobrancas(nome, pares, comControle) {
  const LL = [{ cells: CAB.map(x => t(x, S.hdr)), h:32 }];
  pares.forEach(([a, o], k) => {
    const z = k % 2 ? 1 : 0;
    const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda };
    const traco = t('—', st.num);
    const saldo = o.fantasma ? null : Math.max(0, o.liq - o.baixado);
    LL.push({ cells: [
      t(CODE[a.u] || a.u, st.txt), t(nomeProprio(a.aluno), st.txt), t(a.ra, st.txt),
      t(a.serie, st.txt), t(a.sit, st.txt), t(o.grupo, st.txt), t(o.servico, st.txt),
      n(o.parcela, st.num), n(a.parcelasEsperadas, st.num), t(o.venc || '', st.num),
      o.fantasma ? traco : m(o.orig, st.moeda),
      o.fantasma ? traco : m(o.bolsa, st.moeda),
      o.fantasma ? traco : m(o.liq, st.moeda),
      o.fantasma ? traco : m(o.baixado, st.moeda),
      saldo == null ? traco : m(saldo, st.moeda),
      t(o.dt || '', st.num), t(o.st, st.num), t(o.bolsas, st.txt),
      t(o.fantasma ? '' : o.ref, st.num),
      t(o.sitCob, st.txt), t(o.acao, st.txt),
      o.cancelar ? m(o.cancelar, st.moeda) : traco,
      n(a.mens, st.num), t('', comControle ? S.acao : st.num), t('', st.txt)
    ] });
  });
  const aba = { nome, cols:COLS, linhas:LL, congelar:1, filtro:'A1:' + colName(CAB.length - 1) + LL.length };
  if (comControle) {
    aba.validacao = { ref:'X2:X' + LL.length, opcoes:['Sim','Não','Em análise'],
      dica:'A unidade marca aqui depois de ajustar a cobrança no Totvs.' };
    aba.formatacao = { ref:'X2:X' + LL.length, igual:'Sim' };
  }
  return aba;
}

const todos = [];
ordenados.forEach(a => a.lanc.forEach(o => todos.push([a, o])));
const aCancelar = todos.filter(([, o]) => o.cancelar);
const emDobro = aCancelar.filter(([, o]) => /entrada em dobro/i.test(o.acao));
const saldos   = aCancelar.filter(([, o]) => /saldo/i.test(o.acao));
const totalCancelar = aCancelar.reduce((s, [, o]) => s + o.cancelar, 0);

/* ---------------- fora do padrão ---------------- */
const CAB3 = ['Unidade','Aluno','RA','Situação do aluno','Entrada cobrada','Entrada paga','Saldo em aberto',
  'Parcelas da entrada','Previstas','Mensalidades','Fora do padrão','Resolvido?','Observação da unidade'];
const L3 = [{ cells: CAB3.map(x => t(x, S.hdr)), h:32 }];
ordenados.filter(a => a.fora.length).forEach((a, k) => {
  const z = k % 2 ? 1 : 0;
  const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda };
  const g = a.principal;
  L3.push({ cells: [
    t(CODE[a.u] || a.u, st.txt), t(nomeProprio(a.aluno), st.txt), t(a.ra, st.txt), t(a.sit, st.txt),
    g ? m(g.cobrado, st.moeda) : t('—', st.num),
    g ? m(g.pago, st.moeda) : t('—', st.num),
    g ? m(Math.max(0, g.emAberto), st.moeda) : t('—', st.num),
    n(g ? g.ls.length : 0, st.num), n(a.parcelasEsperadas, st.num), n(a.mens, st.num),
    t(a.fora.join(' · '), st.txt), t('', S.acao), t('', st.txt)
  ] });
});

/* ---------------- resumo ---------------- */
const CAB4 = ['Unidade','Alunos','Entrada em dobro','Saldo de cota emitida cheia','Total a cancelar',
  'Zerada com bolsa','Já cancelada','Sem lançamento','Fora do padrão'];
const L4 = [{ cells: CAB4.map(x => t(x, S.hdr)), h:32 }];
const conta = (as, f) => as.reduce((s, a) => s + a.lanc.filter(f).length, 0);
const grupos = ORDER.map(c => ordenados.filter(a => a.u === c)).filter(as => as.length);
grupos.forEach((as, k) => {
  const z = k % 2 ? 1 : 0;
  const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda };
  const dob = conta(as, o => o.cancelar && /entrada em dobro/i.test(o.acao));
  const sal = conta(as, o => o.cancelar && /saldo/i.test(o.acao));
  const val = as.reduce((s, a) => s + a.lanc.reduce((x, o) => x + (o.cancelar || 0), 0), 0);
  L4.push({ cells: [ t(CODE[as[0].u] || as[0].u, st.txt), n(as.length, st.num),
    n(dob, st.num), n(sal, st.num), m(val, st.moeda),
    n(conta(as, o => o.sitCob === SIT.bolsa), st.num),
    n(conta(as, o => o.sitCob === SIT.cancelada), st.num),
    n(conta(as, o => o.sitCob === SIT.fantasma), st.num),
    n(as.filter(a => a.fora.length).length, st.num) ] });
});
L4.push({ cells: [ t('Rede', S.verde), n(ordenados.length, S.verde),
  n(emDobro.length, S.verde), n(saldos.length, S.verde), t(brl(totalCancelar), S.verde),
  n(conta(ordenados, o => o.sitCob === SIT.bolsa), S.verde),
  n(conta(ordenados, o => o.sitCob === SIT.cancelada), S.verde),
  n(conta(ordenados, o => o.sitCob === SIT.fantasma), S.verde),
  n(ordenados.filter(a => a.fora.length).length, S.verde) ] });
L4.push({ cells: [] });
[
  ['O padrão e como cada cobrança foi lida', S.titulo],
  ['Padrão: cota 1 de R$ 300 e 12 mensalidades. Na ficha financeira a 1ª Cota de Mensalidade é a PARCELA 0 do contrato e as mensalidades são as parcelas 1 a 12 — a cota não é uma das doze. QTD_PARC_CONTRATO = 13 confirma.', S.legenda],
  ['Entrada partida: QTD_PARC_CONTRATO menos 12 diz em quantas parcelas a entrada foi dividida (14 = 2 parcelas, 18 = 6). Parcelas da mesma entrada não são cobrança em dobro.', S.legenda],
  ['Entrada em dobro — o aluno tem a reserva de vaga do bolsão E a 1ª cota do contrato, as duas vivas. Cancela-se a que não recebeu dinheiro.', S.legenda],
  ['Saldo de cota emitida cheia — a 1ª cota saiu pelo valor da tabela, os R$ 300 do bolsão foram baixados por dentro dela e o saldo ficou aberto. Esse saldo não é devido pelo padrão de R$ 300.', S.legenda],
  ['Zerada com bolsa e baixada — STATUS_BOLETO "Baixado" com VALOR_BAIXADO zero. A unidade já resolveu.', S.legenda],
  ['Sem lançamento financeiro — VALOR_ORIGINAL vazio e "SEM BOLETO". O contrato existe e a cobrança não foi gerada: falta lançar, não há o que cancelar.', S.legenda],
  ['Apurado na ficha financeira de 09/09/2026. Alunos de teste ficaram de fora.', S.legenda]
].forEach(([txt, s]) => L4.push({ cells:[t(txt, s)], h: s === S.titulo ? 22 : undefined }));

/* ---------------- grava ---------------- */
if (require.main === module) {
  const buf = pasta([
    abaCobrancas('A ajustar', aCancelar, true),
    abaCobrancas('Todas as cobranças', todos, false),
    { nome:'Fora do padrão', cols:[17,30,13,24,16,14,15,18,11,13,60,12,30], linhas:L3, congelar:1,
      filtro:'A1:' + colName(CAB3.length - 1) + L3.length,
      validacao:{ ref:'L2:L' + L3.length, opcoes:['Sim','Não','Em análise'],
        dica:'A unidade marca aqui quando a estrutura estiver corrigida.' },
      formatacao:{ ref:'L2:L' + L3.length, igual:'Sim' } },
    { nome:'Resumo', cols:[22,10,18,28,17,18,14,16,15], linhas:L4, congelar:1 }
  ]);
  const destino = path.join(SAIDA, 'Entrada da captação 2027 - o que ajustar.xlsx');
  fs.writeFileSync(destino, buf);

  console.log('->', destino, Math.round(buf.length / 1024) + ' KB');
  console.log(ordenados.length + ' alunos ·', todos.length, 'lançamentos de entrada');
  console.log('');
  const porSit = {};
  todos.forEach(([, o]) => { const k = o.grupo + ' · ' + o.sitCob; porSit[k] = (porSit[k] || 0) + 1; });
  Object.entries(porSit).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log('  ' + String(v).padStart(4) + '  ' + k));
  console.log('');
  console.log('A AJUSTAR: ' + aCancelar.length + ' lançamentos · ' + brl(totalCancelar));
  console.log('  entrada em dobro:            ' + String(emDobro.length).padStart(3) + '  ' +
    brl(emDobro.reduce((s, [, o]) => s + o.cancelar, 0)));
  console.log('  saldo de cota emitida cheia: ' + String(saldos.length).padStart(3) + '  ' +
    brl(saldos.reduce((s, [, o]) => s + o.cancelar, 0)));
  console.log('');
  grupos.forEach(as => {
    const c = as.reduce((s, a) => s + a.lanc.filter(o => o.cancelar).length, 0);
    const v = as.reduce((s, a) => s + a.lanc.reduce((x, o) => x + (o.cancelar || 0), 0), 0);
    if (c) console.log('  ' + (CODE[as[0].u] || as[0].u).padEnd(20) + String(c).padStart(3) + '  ' + brl(v));
  });
  console.log('');
  console.log('fora do padrão: ' + ordenados.filter(a => a.fora.length).length + ' alunos');
}

module.exports = { alunos, todos, aCancelar, SIT, CODE, ORDER };
