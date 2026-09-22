/* Os alunos da captação 2027 com duas primeiras cotas.
 *
 *   node motor/duas-cotas.js [pasta-de-saida]
 *
 * O padrão é 1 cota e 12 mensalidades. Quem tem duas cobranças de entrada
 * aparece aqui com uma linha para cada uma, uma embaixo da outra, na mesma
 * faixa de cor — para a dobra ficar visível sem precisar cruzar nada.
 *
 * Duas cobranças de entrada = a reserva de vaga do bolsão E a 1ª Cota de
 * Mensalidade gerada pelo contrato. Não conta como dobra:
 *   · a entrada partida em parcelas (mesma cobrança, vencimentos diferentes);
 *   · o lançamento já cancelado;
 *   · a linha sem lançamento financeiro, que vem do export com VALOR_ORIGINAL
 *     vazio, REF_FINANCEIRO 4837 e VALOR_BAIXADO 4838 em todas — número de
 *     controle do relatório, não cobrança.
 *
 * A 1ª Cota de Mensalidade não é uma das 12: ela é a PARCELA 0 do contrato e as
 * mensalidades são as parcelas 1 a 12. Excluir a cota deixa as 12 de pé. */

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');
const { pasta, colName } = require('./escrever-xlsx.js');

const B = 'C:/Users/ti/Desktop/Matrículas 2027/';
const P_FICHA = process.env.FICHA || (B + '14 de setembro/Ficha financeira detalhada.XLSX');
const SAIDA = process.argv[2] || B;

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

const num = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; };
const ehData = v => /^\d{2}\/\d{2}\/\d{4}$/.test(String(v || '').trim());
const TESTE = /\bteste?s?\b/i;
const PARTICULA = new Set(['de','da','do','das','dos','e','di','du','del','la','van','von','y']);
const nomeProprio = s => String(s || '').trim().toLowerCase().split(/\s+/)
  .map((p, i) => (i > 0 && PARTICULA.has(p)) ? p : p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
const brl = v => 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits:2, maximumFractionDigits:2 });

const RESERVA = /Reserva de vaga/i, COTA = /1ª Cota/i, MENS = /(^|- )Mensalidade$/i;

/* ---------------- leitura ---------------- */
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
  if (!alunos.has(r[3])) alunos.set(r[3], { ra: r[3], aluno: r[4], u: unidade(r[2]),
    sit: r[21] || '', serie: r[17] || '', cobrancas: [], mens: 0 });
  const a = alunos.get(r[3]);
  const s = String(r[34] || '');
  if (MENS.test(s)) { if (String(r[47] || '') !== 'Cancelado') a.mens++; return; }
  if (!RESERVA.test(s) && !COTA.test(s)) return;
  if (fantasma(r) || String(r[47] || '') === 'Cancelado') return;   // não é cobrança viva
  const o = {
    cobranca: RESERVA.test(s) ? 'Reserva de vaga' : '1ª cota do contrato',
    servico: s, parcela: num(r[28]), venc: String(r[41] || '').trim(),
    valor: num(r[42]), bolsa: num(r[43]), baixado: num(r[44]),
    dt: ehData(r[45]) ? String(r[45]).trim() : null,
    st: String(r[47] || ''), bolsas: String(r[48] || '').trim(), ref: String(r[35] || '').trim()
  };
  o.aPagar = o.valor - o.bolsa;
  const falta = Math.max(0, o.aPagar - o.baixado);
  o.situacao = o.baixado > 0 ? (falta < 0.5 ? 'Paga' : 'Paga em parte')
    : o.st === 'Baixado' ? 'Zerada com bolsa e baixada' : 'Em aberto';
  // Parcela baixada não cobra mais nada, mesmo quando a baixa foi por um valor
  // menor. Mostrar saldo numa linha baixada faria a pessoa procurar um boleto
  // que não existe.
  o.emAberto = o.st === 'Baixado' ? 0 : falta;
  a.cobrancas.push(o);
});

/* ---------------- quem tem duas ---------------- */
// Vencimentos diferentes na mesma cobrança são parcelas dela, não dobra.
const doisGrupos = a => new Set(a.cobrancas.map(o => o.cobranca)).size > 1;
const comDuas = [...alunos.values()].filter(doisGrupos).sort((a, b) =>
  ORDER.indexOf(a.u) - ORDER.indexOf(b.u) || String(a.aluno).localeCompare(String(b.aluno), 'pt'));

comDuas.forEach(a => {
  const grupos = ['Reserva de vaga', '1ª cota do contrato'].map(nome => {
    const ls = a.cobrancas.filter(o => o.cobranca === nome);
    return { nome, ls, recebeu: ls.reduce((s, o) => s + o.baixado, 0) };
  }).filter(g => g.ls.length);

  // Fica de pé a cobrança que recebeu dinheiro. Sem dinheiro em nenhuma, fica a
  // reserva de vaga, que é a cobrança certa dos R$ 300 do bolsão.
  const comDinheiro = grupos.filter(g => g.recebeu > 0.005);
  const manter = comDinheiro.length === 1 ? comDinheiro[0]
    : comDinheiro.length > 1 ? null
    : grupos.find(g => g.nome === 'Reserva de vaga') || grupos[0];

  grupos.forEach(g => g.ls.forEach(o => {
    if (manter === null) {
      o.fazer = 'CONFERIR';
      o.porque = 'As duas cobranças receberam pagamento: devolver ou creditar antes de excluir';
      return;
    }
    if (g === manter) {
      o.fazer = 'MANTER';
      o.porque = g.recebeu > 0.005 ? 'É a cobrança que recebeu ' + brl(g.recebeu)
        : 'É a cobrança certa da entrada; nenhuma das duas foi paga';
      return;
    }
    o.fazer = 'EXCLUIR';
    o.excluir = o.emAberto;
    o.porque = o.situacao === 'Zerada com bolsa e baixada'
      ? 'Cobrança em dobro, já zerada com bolsa: estornar a baixa antes de excluir'
      : 'Cobrança em dobro: a entrada já está na outra linha deste aluno';
  }));
  a.grupos = grupos;
});

/* ---------------- células ---------------- */
const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9,
            legenda:10, verde:11, lar:12, larData:13, larMoeda:14, larNum:15 };
const t = (v, s) => ({ v, t:'s', s: s == null ? S.txt : s });
const n = (v, s) => ({ v, t:'n', s: s == null ? S.num : s });
const m = (v, s) => ({ v, t:'m', s: s == null ? S.moeda : s });

const CAB = ['Unidade','Aluno','RA','Série','Situação do aluno','O QUE FAZER','Por quê',
  'Cobrança','Serviço no Totvs','Parcela','Vencimento','Valor da cota','Bolsa aplicada',
  'Valor a pagar','Valor baixado','Em aberto','Situação da cobrança','Data da baixa',
  'Bolsas cadastradas','Ref. financeiro','Mensalidades','Excluído?'];
const COLS = [17,30,13,9,24,13,52,20,34,8,12,13,14,13,13,13,24,12,34,13,13,12];

/* Resolvido é quando a dobra já foi tratada pela unidade: a cobrança que sobra
   está zerada com bolsa e baixada. O que ainda tem cobrança viva, ou dinheiro
   dos dois lados, é para resolver. */
const resolvido = a => a.cobrancas.filter(o => o.fazer === 'EXCLUIR').length &&
  a.cobrancas.filter(o => o.fazer === 'EXCLUIR').every(o => o.situacao === 'Zerada com bolsa e baixada');
const paraResolver = comDuas.filter(a => !resolvido(a));
const jaResolvidos = comDuas.filter(resolvido);

function abaAlunos(nome, lista, comControle) {
  const LL = [{ cells: CAB.map(x => t(x, S.hdr)), h:32 }];
  let bloco = 0;
  lista.forEach(a => {
    // A faixa de cor muda por aluno, não por linha: as duas cobranças do mesmo
    // aluno ficam no mesmo bloco e a dobra se lê de relance.
    const z = bloco++ % 2;
    const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda };
    const ordem = { 'Reserva de vaga':0, '1ª cota do contrato':1 };
    a.cobrancas.slice()
      .sort((x, y) => ordem[x.cobranca] - ordem[y.cobranca] || x.parcela - y.parcela)
      .forEach(o => {
        LL.push({ cells: [
          t(CODE[a.u] || a.u, st.txt), t(nomeProprio(a.aluno), st.txt), t(a.ra, st.txt),
          t(a.serie, st.txt), t(a.sit, st.txt),
          t(o.fazer, o.fazer === 'MANTER' ? S.acao : st.num), t(o.porque, st.txt),
          t(o.cobranca, st.txt), t(o.servico, st.txt), n(o.parcela, st.num), t(o.venc, st.num),
          m(o.valor, st.moeda), m(o.bolsa, st.moeda), m(o.aPagar, st.moeda),
          m(o.baixado, st.moeda), m(o.emAberto, st.moeda),
          t(o.situacao, st.txt), t(o.dt || '', st.num), t(o.bolsas, st.txt), t(o.ref, st.num),
          n(a.mens, st.num), t('', comControle ? S.acao : st.num)
        ] });
      });
  });
  const aba = { nome, cols:COLS, linhas:LL, congelar:1, filtro:'A1:' + colName(CAB.length - 1) + LL.length };
  if (comControle) {
    aba.validacao = { ref:'V2:V' + LL.length, opcoes:['Sim','Não','Em análise'],
      dica:'Marque Sim depois de excluir a cobrança no Totvs.' };
    aba.formatacao = { ref:'V2:V' + LL.length, igual:'Sim' };
  }
  return aba;
}

/* ---------------- resumo ---------------- */
const CAB2 = ['Unidade','Para resolver','Cobranças a excluir','Valor em aberto','Conferir (as duas pagas)','Já resolvidos'];
const L2 = [{ cells: CAB2.map(x => t(x, S.hdr)), h:32 }];
const grupos = ORDER.map(c => comDuas.filter(a => a.u === c)).filter(as => as.length);
const linhaResumo = (rot, as, s) => {
  const pend = as.filter(a => !resolvido(a));
  const ex = pend.flatMap(a => a.cobrancas.filter(o => o.fazer === 'EXCLUIR'));
  return [ t(rot, s.txt), n(pend.length, s.num), n(ex.length, s.num),
    s.moeda ? m(ex.reduce((x, o) => x + o.emAberto, 0), s.moeda) : t(brl(ex.reduce((x, o) => x + o.emAberto, 0)), s.txt),
    n(pend.filter(a => a.cobrancas.some(o => o.fazer === 'CONFERIR')).length, s.num),
    n(as.filter(resolvido).length, s.num) ];
};
grupos.forEach((as, k) => {
  const z = k % 2 ? 1 : 0;
  L2.push({ cells: linhaResumo(CODE[as[0].u] || as[0].u, as,
    { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num, moeda: z ? S.larMoeda : S.moeda }) });
});
L2.push({ cells: linhaResumo('Rede', comDuas, { txt: S.verde, num: S.verde, moeda: null }) });
L2.push({ cells: [] });
const dataFicha = (P_FICHA.match(/(\d{2}) de (\w+)/) || [])[0] || '';
[
  ['Como ler', S.titulo],
  ['O padrão é 1 cota de entrada e 12 mensalidades. Cada aluno aparece em duas linhas, na mesma faixa de cor: são as duas cobranças de entrada que ele tem no Totvs.', S.legenda],
  ['MANTER é a cobrança que recebeu o pagamento. EXCLUIR é a que sobra. Quando nenhuma das duas foi paga, fica de pé a reserva de vaga, que é a cobrança certa dos R$ 300 do bolsão.', S.legenda],
  ['CONFERIR é quando as duas cobranças receberam dinheiro: aí não é exclusão, é devolução ou crédito.', S.legenda],
  ['"Já resolvidos" são os alunos em que a cobrança sobrando foi zerada com bolsa de 100% e baixada sem valor. Estão na segunda aba só para registro; não precisam de ação.', S.legenda],
  ['A 1ª Cota de Mensalidade não é uma das 12 mensalidades: no Totvs ela é a PARCELA 0 do contrato, e as mensalidades são as parcelas 1 a 12. Excluir a cota deixa as 12 de pé.', S.legenda],
  ['Lançamentos cancelados e linhas sem lançamento financeiro ficaram de fora. Ficha financeira de ' + dataFicha + ', alunos de teste excluídos.', S.legenda]
].forEach(([txt, s]) => L2.push({ cells:[t(txt, s)], h: s === S.titulo ? 22 : undefined }));

/* ---------------- grava ---------------- */
if (require.main === module) {
  const abaPend = abaAlunos('Para resolver', paraResolver, true);
  const buf = pasta([
    abaPend,
    abaAlunos('Já resolvidos', jaResolvidos, false),
    { nome:'Resumo', cols:[22,16,20,18,24,16], linhas:L2, congelar:1 }
  ]);
  const destino = path.join(SAIDA, 'Alunos com 2 primeiras cotas - para resolver.xlsx');
  fs.writeFileSync(destino, buf);

  console.log('->', destino, Math.round(buf.length / 1024) + ' KB');
  console.log(comDuas.length + ' alunos com 2 cotas: ' + paraResolver.length + ' para resolver · ' +
    jaResolvidos.length + ' já resolvidos');
  const ex = paraResolver.flatMap(a => a.cobrancas.filter(o => o.fazer === 'EXCLUIR'));
  console.log('  a excluir: ' + ex.length + ' cobranças · ' + brl(ex.reduce((s, o) => s + o.emAberto, 0)) + ' em aberto');
  console.log('  conferir: ' + paraResolver.filter(a => a.cobrancas.some(o => o.fazer === 'CONFERIR')).length + ' alunos');
  console.log('');
  grupos.forEach(as => {
    const pend = as.filter(a => !resolvido(a));
    if (!pend.length) return;
    const e = pend.flatMap(a => a.cobrancas.filter(o => o.fazer === 'EXCLUIR'));
    console.log('  ' + (CODE[as[0].u] || as[0].u).padEnd(20) + String(pend.length).padStart(3) + ' para resolver · ' +
      brl(e.reduce((s, o) => s + o.emAberto, 0)));
  });
}

module.exports = { alunos, comDuas, paraResolver, jaResolvidos, CODE, ORDER };
