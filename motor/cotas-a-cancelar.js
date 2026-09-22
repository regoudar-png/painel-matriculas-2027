/* Todo lançamento de entrada da captação 2027, um por linha, com o veredito.
 *
 *   node motor/cotas-a-cancelar.js [pasta-de-saida]
 *
 * O desenho correto é uma entrada de R$ 300 e 12 mensalidades. Quando o
 * contrato é gerado, o Totvs lança a "1ª Cota de Mensalidade" por cima da
 * reserva de vaga já paga no bolsão — é essa cota que sobra.
 *
 * A leitura direta da ficha financeira mostra que boa parte já foi resolvida
 * pelas unidades, de duas maneiras que não apareciam no painel:
 *   · cancelamento do lançamento (STATUS_BOLETO = Cancelado);
 *   · bolsa de 100% e baixa da cota a zero (Baixado com VALOR_BAIXADO = 0).
 * E há um terceiro grupo que não é cobrança nenhuma: linhas sem lançamento
 * financeiro, que vêm do export com VALOR_ORIGINAL vazio, REF_FINANCEIRO 4837
 * e VALOR_BAIXADO 4838 em todas — número de controle do relatório, não dinheiro.
 * Contar essas três coisas como cobrança em dobro é o que inflava a lista. */

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');
const { pasta, colName } = require('./escrever-xlsx.js');

const B = 'C:/Users/ti/Desktop/Matrículas 2027/';
const P_FICHA = B + '09 de setembro/Ficha financeira.XLSX';
const SAIDA = process.argv[2] || B;

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

const RESERVA = /Reserva de vaga/i, COTA = /1ª Cota/i, MENS = /(^|- )Mensalidade$/i;
const ENTRADA_ESPERADA = 300;

const L = lerPlanilha(P_FICHA).abas.filter(a => a.linhas)
  .sort((a, b) => b.linhas.length - a.linhas.length)[0].linhas;
const cap = L.slice(1).filter(r => r && r[3] && r[19] === 'MATRÍCULA' && !TESTE.test(String(r[4] || '')));

// Sem VALOR_ORIGINAL e sem boleto: o contrato existe e o financeiro não foi
// gerado. Não é cobrança, então não entra na conta de duplicidade.
const semLancamento = r => num(r[42]) < 0.005 && String(r[47] || '') === 'SEM BOLETO';

// A mesma parcela vem duas vezes quando o boleto tem dois status de remessa.
const vistos = new Set();
const linhas = cap.filter(r => {
  const ref = String(r[35] || '').trim();
  if (!ref || semLancamento(r)) return true;
  const k = r[3] + '|' + ref;
  if (vistos.has(k)) return false;
  vistos.add(k); return true;
});

const alunos = new Map();
linhas.forEach(r => {
  if (!alunos.has(r[3])) alunos.set(r[3], {
    ra: r[3], aluno: r[4], u: unidade(r[2]), sit: r[21] || '', serie: r[17] || '',
    plano: r[25] || '', contrato: r[26] || '', entradas: [], mens: 0, mensCanc: 0 });
  const a = alunos.get(r[3]);
  const s = String(r[34] || '');
  const cancelado = String(r[47] || '') === 'Cancelado';
  if (MENS.test(s)) { cancelado ? a.mensCanc++ : a.mens++; return; }
  if (!RESERVA.test(s) && !COTA.test(s)) return;
  const o = {
    tipo: RESERVA.test(s) ? 'Reserva de vaga' : '1ª cota do contrato',
    servico: s, venc: String(r[41] || '').trim(), emissao: String(r[40] || '').trim(),
    orig: num(r[42]), bolsa: num(r[43]), baixado: num(r[44]),
    dt: ehData(r[45]) ? String(r[45]).trim() : null,
    st: String(r[47] || ''), bolsas: String(r[48] || '').trim(),
    ref: String(r[35] || '').trim(), semLanc: semLancamento(r)
  };
  o.liq = o.orig - o.bolsa;
  a.entradas.push(o);
});

/* ---------------- veredito de cada lançamento ---------------- */
const VER = {
  semLanc:  'Sem lançamento financeiro',
  cancelada:'Já cancelada',
  paga:     'Paga',
  bolsa:    'Baixada com bolsa de 100%',
  aberta:   'Em aberto'
};
const vereditoDe = o => o.semLanc ? VER.semLanc
  : o.st === 'Cancelado' ? VER.cancelada
  : o.baixado > 0 ? VER.paga
  : o.st === 'Baixado' ? VER.bolsa
  : VER.aberta;

alunos.forEach(a => {
  a.entradas.forEach(o => { o.veredito = vereditoDe(o); });
  const quitadas = a.entradas.filter(o => o.veredito === VER.paga || o.veredito === VER.bolsa);
  const abertas  = a.entradas.filter(o => o.veredito === VER.aberta);
  a.quitadas = quitadas; a.abertas = abertas;
  a.pagasEmDinheiro = a.entradas.filter(o => o.veredito === VER.paga);
  const manter = abertas.slice().sort((x, y) =>
    (x.tipo === 'Reserva de vaga' ? 0 : 1) - (y.tipo === 'Reserva de vaga' ? 0 : 1) ||
    x.liq - y.liq)[0];

  a.entradas.forEach(o => {
    if (o.veredito === VER.semLanc) { o.acao = 'Gerar o financeiro: o contrato existe e a cobrança não'; return; }
    if (o.veredito === VER.cancelada) { o.acao = 'Nada a fazer'; return; }
    if (o.veredito === VER.bolsa) { o.acao = 'Nada a fazer: a unidade zerou com bolsa e baixou'; return; }
    if (o.veredito === VER.paga) {
      o.acao = a.pagasEmDinheiro.length > 1
        ? 'Pagamento em dobro: devolver ou creditar' : 'Nada a fazer';
      return;
    }
    // Em aberto: só é para cancelar quando outra entrada já está resolvida.
    if (quitadas.length) { o.acao = 'CANCELAR'; o.cancelar = true; return; }
    // Nenhuma paga e mais de uma aberta: uma tem de ficar de pé para a família
    // pagar. Guardo a reserva de vaga, que é a cobrança certa; sem reserva,
    // guardo a de menor valor, que é a que mais se aproxima dos R$ 300.
    if (abertas.length > 1) {
      if (o !== manter) { o.acao = 'CANCELAR (são ' + abertas.length + ' abertas; fica de pé a de ' +
        brl(manter.liq) + ')'; o.cancelar = true; }
      else o.acao = 'Manter esta e cancelar as outras ' + (abertas.length - 1) + ' · entrada ainda não paga';
      return;
    }
    o.acao = 'Entrada ainda não paga: cobrar ou conferir pagamento na Layers';
  });

  /* estrutura do aluno em relação ao padrão: 1 entrada de R$300 + 12 mensalidades */
  const fora = [];
  // Cobrança viva é a que ainda representa dinheiro: paga ou em aberto. A
  // cancelada e a baixada com bolsa já foram resolvidas pela unidade, e a sem
  // lançamento nunca existiu — nenhuma das três é desvio do padrão.
  const viva = a.entradas.filter(o => o.veredito === VER.paga || o.veredito === VER.aberta);
  if (!viva.length) fora.push('sem nenhuma cobrança de entrada');
  else if (viva.length > 1) fora.push(viva.length + ' cobranças de entrada');
  const res = viva.filter(o => o.tipo === 'Reserva de vaga');
  if (res.length && res.some(o => Math.abs(o.orig - ENTRADA_ESPERADA) > 0.005))
    fora.push('reserva fora dos R$ 300 (' + res.map(o => 'R$ ' + o.orig.toFixed(2)).join(', ') + ')');
  // Sem contrato fechado não se espera plano do ano: quem ainda está só com a
  // reserva de vaga não entra nessa conta.
  const esperaPlano = /^(Matriculado|Pré-Matriculado)$/.test(a.sit);
  if (a.mens === 0) { if (esperaPlano) fora.push('sem mensalidade lançada'); }
  else if (a.mens > 12) fora.push(a.mens + ' mensalidades: plano do ano em dobro');
  else if (a.mens < 12) fora.push('só ' + a.mens + ' mensalidades');
  if (a.pagasEmDinheiro.length > 1) fora.push('pagou a entrada duas vezes');
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

/* ---------------- aba de cobranças (o modelo) ---------------- */
const CAB = ['Unidade','Aluno','RA','Série','Situação','Cobrança','Serviço no Totvs','Vencimento',
  'Valor original','Bolsa / deduções','Valor líquido','Valor baixado','Data da baixa',
  'Status do boleto','Bolsas cadastradas','Ref. financeiro','Situação da cobrança','O que fazer',
  'Mensalidades','Resolvido?','Observação da unidade'];

function abaCobrancas(nome, pares, comControle) {
  const L = [{ cells: CAB.map(x => t(x, S.hdr)), h:30 }];
  pares.forEach(([a, o], k) => {
    const z = k % 2 ? 1 : 0;
    const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda };
    L.push({ cells: [
      t(CODE[a.u] || a.u, st.txt), t(nomeProprio(a.aluno), st.txt), t(a.ra, st.txt),
      t(a.serie, st.txt), t(a.sit, st.txt), t(o.tipo, st.txt), t(o.servico, st.txt),
      t(o.venc || '', st.num),
      o.semLanc ? t('—', st.num) : m(o.orig, st.moeda),
      o.semLanc ? t('—', st.num) : m(o.bolsa, st.moeda),
      o.semLanc ? t('—', st.num) : m(o.liq, st.moeda),
      o.semLanc ? t('—', st.num) : m(o.baixado, st.moeda),
      t(o.dt || '', st.num), t(o.st, st.num), t(o.bolsas, st.txt),
      t(o.semLanc ? '' : o.ref, st.num),
      t(o.veredito, st.txt), t(o.acao, st.txt), n(a.mens, st.num),
      t('', comControle ? S.acao : st.num), t('', st.txt)
    ] });
  });
  const aba = { nome, cols:[17,30,13,9,22,19,34,12,13,14,13,13,12,13,34,12,24,46,12,12,30],
    linhas:L, congelar:1, filtro:'A1:' + colName(CAB.length - 1) + L.length };
  if (comControle) {
    aba.validacao = { ref:'T2:T' + L.length, opcoes:['Sim','Não','Em análise'],
      dica:'A unidade marca aqui depois de cancelar a cobrança no Totvs.' };
    aba.formatacao = { ref:'T2:T' + L.length, igual:'Sim' };
  }
  return aba;
}

const todos = [];
ordenados.forEach(a => a.entradas.forEach(o => todos.push([a, o])));
const aCancelar = todos.filter(([, o]) => o.cancelar);

/* ---------------- aba fora do padrão ---------------- */
const CAB3 = ['Unidade','Aluno','RA','Situação','Cobranças de entrada vivas','Entrada paga',
  'Mensalidades','Plano de pagamento','Fora do padrão','Resolvido?','Observação da unidade'];
const L3 = [{ cells: CAB3.map(x => t(x, S.hdr)), h:30 }];
ordenados.filter(a => a.fora.length).forEach((a, k) => {
  const z = k % 2 ? 1 : 0;
  const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda };
  const viva = a.entradas.filter(o => o.veredito !== VER.cancelada && o.veredito !== VER.semLanc);
  L3.push({ cells: [
    t(CODE[a.u] || a.u, st.txt), t(nomeProprio(a.aluno), st.txt), t(a.ra, st.txt), t(a.sit, st.txt),
    n(viva.length, st.num),
    m(a.pagasEmDinheiro.reduce((s, o) => s + o.baixado, 0), st.moeda),
    n(a.mens, st.num), t(a.plano, st.txt),
    t(a.fora.join(' · '), st.txt), t('', S.acao), t('', st.txt)
  ] });
});

/* ---------------- aba resumo ---------------- */
const CAB4 = ['Unidade','Alunos','1ª cota a cancelar','Valor a cancelar','Já cancelada',
  'Baixada com bolsa','Paga','Sem lançamento financeiro','Fora do padrão'];
const L4 = [{ cells: CAB4.map(x => t(x, S.hdr)), h:30 }];
const conta = (as, f) => as.reduce((s, a) => s + a.entradas.filter(f).length, 0);
const linhasResumo = ORDER.map(c => ordenados.filter(a => a.u === c)).filter(as => as.length);
linhasResumo.forEach((as, k) => {
  const z = k % 2 ? 1 : 0;
  const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda };
  const canc = as.reduce((s, a) => s + a.entradas.filter(o => o.cancelar).length, 0);
  const val = as.reduce((s, a) => s + a.entradas.filter(o => o.cancelar).reduce((x, o) => x + o.liq, 0), 0);
  L4.push({ cells: [
    t(CODE[as[0].u] || as[0].u, st.txt), n(as.length, st.num), n(canc, st.num), m(val, st.moeda),
    n(conta(as, o => o.veredito === VER.cancelada), st.num),
    n(conta(as, o => o.veredito === VER.bolsa), st.num),
    n(conta(as, o => o.veredito === VER.paga), st.num),
    n(conta(as, o => o.veredito === VER.semLanc), st.num),
    n(as.filter(a => a.fora.length).length, st.num)
  ] });
});
const totalCancelar = aCancelar.reduce((s, [, o]) => s + o.liq, 0);
L4.push({ cells: [ t('Rede', S.verde), n(ordenados.length, S.verde), n(aCancelar.length, S.verde),
  t(brl(totalCancelar), S.verde),
  n(conta(ordenados, o => o.veredito === VER.cancelada), S.verde),
  n(conta(ordenados, o => o.veredito === VER.bolsa), S.verde),
  n(conta(ordenados, o => o.veredito === VER.paga), S.verde),
  n(conta(ordenados, o => o.veredito === VER.semLanc), S.verde),
  n(ordenados.filter(a => a.fora.length).length, S.verde) ] });
L4.push({ cells: [] });
[
  ['Como cada cobrança foi classificada', S.titulo],
  ['Padrão correto: uma entrada de R$ 300 (reserva de vaga do bolsão) e 12 mensalidades.', S.legenda],
  ['Paga — VALOR_BAIXADO maior que zero com data de baixa. A família pagou.', S.legenda],
  ['Baixada com bolsa de 100% — STATUS_BOLETO "Baixado" com VALOR_BAIXADO zero. A unidade aplicou a bolsa e baixou a cota. Resolvido, nada a cancelar.', S.legenda],
  ['Já cancelada — STATUS_BOLETO "Cancelado". Resolvido.', S.legenda],
  ['Sem lançamento financeiro — VALOR_ORIGINAL vazio e "SEM BOLETO". O contrato existe e a cobrança não foi gerada; não há o que cancelar, há o que lançar.', S.legenda],
  ['Em aberto — cobrança viva. Vira CANCELAR quando o aluno já tem outra entrada paga ou baixada com bolsa; senão é entrada a receber.', S.legenda],
  ['Apurado na ficha financeira de 09/09/2026. Alunos de teste ficaram de fora.', S.legenda]
].forEach(([txt, s]) => L4.push({ cells:[t(txt, s)], h: s === S.titulo ? 22 : undefined }));

/* ---------------- grava ---------------- */
const buf = pasta([
  abaCobrancas('A cancelar', aCancelar, true),
  abaCobrancas('Todas as cobranças', todos, false),
  { nome:'Fora do padrão', cols:[17,30,13,22,15,14,13,30,52,12,30], linhas:L3, congelar:1,
    filtro:'A1:' + colName(CAB3.length - 1) + L3.length,
    validacao:{ ref:'J2:J' + L3.length, opcoes:['Sim','Não','Em análise'],
      dica:'A unidade marca aqui quando a estrutura estiver corrigida.' },
    formatacao:{ ref:'J2:J' + L3.length, igual:'Sim' } },
  { nome:'Resumo', cols:[22,10,20,17,14,19,10,26,15], linhas:L4, congelar:1 }
]);
if (require.main === module) {
  const destino = path.join(SAIDA, '1ª cota - o que cancelar - captação 2027.xlsx');
  fs.writeFileSync(destino, buf);

  /* ---------------- relato ---------------- */
  console.log('->', destino, Math.round(buf.length / 1024) + ' KB');
  console.log(ordenados.length + ' alunos de captação ·', todos.length, 'lançamentos de entrada');
  console.log('');
  const porVer = {};
  todos.forEach(([, o]) => { const k = o.tipo + ' · ' + o.veredito; porVer[k] = (porVer[k] || 0) + 1; });
  Object.entries(porVer).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log('  ' + String(v).padStart(4) + '  ' + k));
  console.log('');
  console.log('A CANCELAR: ' + aCancelar.length + ' lançamentos de 1ª cota · ' + brl(totalCancelar));
  linhasResumo.forEach(as => {
    const c = as.reduce((s, a) => s + a.entradas.filter(o => o.cancelar).length, 0);
    const v = as.reduce((s, a) => s + a.entradas.filter(o => o.cancelar).reduce((x, o) => x + o.liq, 0), 0);
    if (c) console.log('  ' + (CODE[as[0].u] || as[0].u).padEnd(20) + String(c).padStart(3) + '  ' + brl(v));
  });
  console.log('');
  console.log('fora do padrão: ' + ordenados.filter(a => a.fora.length).length + ' alunos');


}

// Requerido como módulo, devolve o retrato para conferências pontuais.
module.exports = { alunos, todos, aCancelar, VER, CODE, ORDER };
