/* Renovação 2027: alunos a preço de tabela cheia ou sem contrato válido.
 *
 *   node motor/bolsa-futura.js [pasta-de-saida]
 *
 * A Bolsa Futura é a proposta de renovação entregue às famílias. Quem renovou
 * sem ela ficou com um contrato de anuidade cheia — R$ 37 a 40 mil contra um
 * ticket mediano de R$ 10 mil. Enquanto o contrato errado está de pé ele entra
 * no ticket médio da rede e o empurra para cima; corrigir faz o ticket cair,
 * que é o número verdadeiro aparecendo.
 *
 * São quatro situações, e a planilha traz uma linha por contrato para que a
 * unidade veja qual cancelar e qual manter:
 *
 *   CONTRATO EM DOBRO   dois contratos ativos, um com bolsa e um sem. O aluno
 *                       escolheu uma trilha (9º militar, 9º vestibular) e
 *                       depois trocou; os dois contratos ficaram de pé.
 *   LANÇAR A BOLSA      um único contrato ativo, a bolsa nunca foi lançada.
 *   REFAZER O CONTRATO  o contrato que tinha a bolsa foi cancelado e nada
 *                       entrou no lugar.
 *   GERAR O CONTRATO    matrícula fechada sem nenhuma mensalidade lançada.
 *
 * VALOR_ORIGINAL é o valor de tabela e BOLSA/DEDUÇÕES o desconto; o que a
 * família paga é a diferença — a mesma conta que o painel usa no ticket. */

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');
const { pasta, colName } = require('./escrever-xlsx.js');

const B = 'C:/Users/ti/Desktop/Matrículas 2027/';
const DIA = process.env.DIA || '20 de outubro';
const P_CAD = process.env.CADASTRO || (B + DIA + '/Ficha cadastral.XLSX');
const P_FIN = process.env.FICHA || (B + DIA + '/Ficha financeira.XLSX');
const SAIDA = process.argv[2] || B;

const CODE = {
  BG:'Bangu', CG:'Campo Grande', CX:'Duque de Caxias', MD:'Madureira', NI:'Nova Iguaçu',
  RM:'Rocha Miranda', RT:'Retiro dos Artistas', SJ:'São João de Meriti', TQ:'Taquara',
  TJ:'Tijuca', AM:'Américas'
};
const ORDER = ['CG','AM','BG','RM','CX','MD','TQ','TJ','SJ','NI','RT'];
const DE_FILIAL = {
  'TAQUARA':'TQ','DUQUE DE CAXIAS':'CX','BANGU':'BG','ROCHA MIRANDA':'RM','NOVA IGUACU':'NI',
  'SAO JOAO DE MERITI':'SJ','CAMPO GRANDE':'CG','RETIRO DOS ARTISTAS':'RT','TIJUCA':'TJ',
  'MADUREIRA':'MD','RECREIO':'AM'
};
const semAcento = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const unidade = f => { const s = semAcento(f).toUpperCase();
  for (const k in DE_FILIAL) if (s.includes(k)) return DE_FILIAL[k]; return '??'; };
const num = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isNaN(x) ? 0 : x; };
const brl = v => 'R$ ' + (Math.round(v * 100) / 100).toLocaleString('pt-BR', { minimumFractionDigits: 2 });
const pc = v => (v * 100).toFixed(1).replace('.', ',') + '%';
const DATA = /^\d{2}\/\d{2}\/\d{4}$/;
const MENS = /Mensalidade/i, COTA = /1ª Cota|Reserva de vaga/i, ANUIDADE = /\bAnuidade\b/i;
const TESTE = /\bteste?s?\b/i;
const AVANCOU = new Set(['Matriculado','Pré-Matriculado','Reserva de vaga - Matriculado','Reserva de vaga - Pré-Mat']);

/* A Bolsa Futura que a unidade propôs à família. Não está em nenhum relatório
   do Totvs: vem da lista do processo de renovação, pelo RA. */
// A proposta fica em dados/bolsa-futura.json, fora do repositório: é RA de
// aluno. Sem o arquivo o script roda igual, só sem a coluna de referência.
const P_PROPOSTA = process.env.PROPOSTA || path.join(__dirname, '..', 'dados', 'bolsa-futura.json');
const PROPOSTA = fs.existsSync(P_PROPOSTA) ? JSON.parse(fs.readFileSync(P_PROPOSTA, 'utf8')) : {};

/* ---------------- leitura ---------------- */

const cad = lerPlanilha(P_CAD).abas['0'].linhas;
const fin = lerPlanilha(P_FIN).abas['0'].linhas;

const alunos = new Map();
for (let i = 1; i < cad.length; i++) {
  const r = cad[i];
  if (String(r[3]) !== '2027' || String(r[10]) !== 'REMATRÍCULA') continue;
  const ra = r[12]; if (!ra || TESTE.test(String(r[13] || ''))) continue;
  if (!alunos.has(ra)) alunos.set(ra, { ra, nome: String(r[13] || '').trim(),
    u: unidade(r[1]), serie: r[5], mats: [], cs: [] });
  alunos.get(ra).mats.push({ turma: r[6], sit: r[9] });
}

for (let i = 1; i < fin.length; i++) {
  const r = fin[i];
  const a = alunos.get(r[3]); if (!a) continue;
  // CODCONTRATO é o número que a unidade procura no Totvs. A coluna TURMA vem
  // vazia neste export; a trilha entra na chave porque o mesmo contrato pode
  // não aparecer e duas trilhas nunca são o mesmo contrato.
  const k = String(r[24] || r[23] || '~') + '|' + r[17];
  let c = a.cs.find(x => x.k === k);
  if (!c) a.cs.push(c = { k, id: r[24] || null, trilha: r[17], sit: r[26] || null,
    curso: String(r[14] || '').trim(),
    sitMat: String(r[21] || '').trim(), planoPg: String(r[25] || '').trim(),
    qtd: num(r[27]) || null, ls: [] });
  c.ls.push(r);
}

for (const a of alunos.values()) for (const c of a.cs) {
  const mens = c.ls.filter(r => +r[28] >= 1 && +r[28] <= 12 && String(r[31]) === '2027'
    && MENS.test(String(r[34])) && !/1ª Cota/i.test(String(r[34])));
  const cotas = c.ls.filter(r => COTA.test(String(r[34])));
  // A família pode quitar o ano inteiro num lançamento só: o serviço "Anuidade".
  // É contrato completo, não plano faltando.
  const anuidade = c.ls.filter(r => ANUIDADE.test(String(r[34])) && num(r[42]) > 0
    && String(r[47] || '') !== 'Cancelado');
  // Lançamento cancelado não é cobrança: num contrato ativo ele é um buraco no
  // plano do ano. É o mesmo critério que o painel usa.
  const vivas = mens.filter(r => String(r[47] || '') !== 'Cancelado');
  c.nMens = vivas.length;
  c.nCanc = mens.length - vivas.length;
  c.avista = anuidade.length > 0;
  c.cheiaMens = vivas.reduce((s, r) => s + num(r[42]), 0);
  c.cheia = c.avista ? anuidade.reduce((s, r) => s + num(r[42]), 0)
    : c.cheiaMens + cotas.reduce((s, r) => s + num(r[42]), 0);
  c.mensal = vivas.length ? num(vivas[0][42]) : 0;
  c.bolsa = c.ls.reduce((s, r) => s + num(r[43]), 0);
  c.pct = c.cheia ? c.bolsa / c.cheia : 0;
  c.nCotas = cotas.filter(r => String(r[47] || '') !== 'Cancelado').length;
  // Contraturno (CCT / Semi-Integral) é serviço à parte, contratado junto com
  // a série. Dois contratos assim não são duas trilhas: um deles é o integral.
  c.complementar = /^CCT$/i.test(c.curso) || /semi-?integral/i.test(String(c.trilha || ''));
  c.fantasma = c.ls.every(r => !r[42] && String(r[47] || '') === 'SEM BOLETO');
  c.ativo = c.sit === 'ATIVO' && !c.fantasma;
  // O complementar não entra na conta da trilha do aluno.
  c.conta = c.ativo && !c.complementar;
  c.plano = c.avista ? 'avista' : c.nMens === 12 ? 'mensal'
    : c.nMens === 0 ? 'vazio' : 'parcial';
  // A anuidade líquida: na à vista é o lançamento único, no mensal são as 12
  // mensalidades líquidas mais a cota de entrada — o critério do painel.
  c.anual = c.avista ? c.cheia - c.bolsa
    : c.nMens === 12 ? c.cheiaMens - vivas.reduce((s, r) => s + num(r[43]), 0) +
      (a.u === 'AM' ? 600 : 300) : null;
  const paga = (c.avista ? anuidade : cotas).find(r => num(r[44]) > 0 && DATA.test(String(r[45] || '')));
  const vivo = (c.avista ? anuidade : cotas).filter(r => String(r[47] || '') !== 'Cancelado');
  c.cobranca = paga ? (c.avista ? 'anuidade paga em ' : 'paga em ') + paga[45]
    : c.avista ? 'anuidade à vista em aberto · vence ' + (vivo[0] && vivo[0][33] || '?')
    : !cotas.length ? 'sem cota lançada'
    : !vivo.length ? 'cota cancelada'
    : vivo.every(r => String(r[47] || '') === 'SEM BOLETO') ? 'sem boleto'
    : 'em aberto · vence ' + (vivo[0][33] || '?');
}

/* ---------------- classificação ---------------- */

// A qualidade de cada contrato ativo. Plano completo é 12 mensalidades ou a
// anuidade à vista; com bolsa, está certo.
const COMPLETO = c => c.plano === 'mensal' || c.plano === 'avista';
const BOM = c => COMPLETO(c) && c.bolsa > 0;
const CHEIO = c => COMPLETO(c) && c.bolsa <= 0;
const VAZIO = c => c.plano === 'vazio';
const PARCIAL = c => c.plano === 'parcial';
// Quanto de bolsa o contrato tem, em % quando dá para calcular.
const bol = c => c.pct > 0 ? pc(c.pct) : brl(c.bolsa);

for (const a of alunos.values()) {
  const vivos = a.cs.filter(c => !c.fantasma);
  const ativos = vivos.filter(c => c.conta);
  const bons = ativos.filter(BOM);
  const sobrando = ativos.filter(c => CHEIO(c) || VAZIO(c));
  const avancou = a.mats.some(m => AVANCOU.has(m.sit));

  if (bons.length && sobrando.length) a.caso = 'dobro';
  // Duas trilhas, as duas com bolsa: não há linha errada para cancelar, há uma
  // escolha a confirmar com a família.
  else if (bons.length > 1) a.caso = 'escolher';
  else if (ativos.some(CHEIO)) a.caso = 'sem-bolsa';
  else if (ativos.length && ativos.every(VAZIO)) a.caso = 'sem-mensalidade';
  else if (ativos.some(PARCIAL)) a.caso = 'parcelas-fora';
  else if (ativos.some(c => c.plano === 'mensal' && !c.nCotas)) a.caso = 'sem-cota';
  else if (!ativos.length && vivos.some(c => c.bolsa > 0)) a.caso = 'cancelado';
  else if (!vivos.length && avancou) a.caso = 'sem-contrato';
  else if (!ativos.length && vivos.length && avancou) a.caso = 'cancelado-sem-bolsa';
  else if (PROPOSTA[a.ra] && !ativos.length) a.caso = 'nao-renovou';
  else a.caso = null;

  // A bolsa de referência: a proposta da unidade, ou a bolsa que o próprio
  // aluno tem em outro contrato.
  const outro = vivos.filter(c => c.bolsa > 0).sort((x, y) => y.pct - x.pct)[0];
  if (PROPOSTA[a.ra]) { a.refPct = PROPOSTA[a.ra][0]; a.pago26 = PROPOSTA[a.ra][1];
    a.refFonte = 'lista da Bolsa Futura'; }
  else if (outro) { a.refPct = outro.pct; a.refFonte = 'outro contrato do aluno'; }
  else { a.refPct = null; a.refFonte = ''; }
}

const CASOS = {
  dobro: { t:'EXCLUIR O CONTRATO', ordem:1, nome:'Contrato em dobro' },
  escolher: { t:'ESCOLHER A TRILHA', ordem:2, nome:'Duas trilhas com contrato' },
  'sem-bolsa': { t:'LANÇAR A BOLSA', ordem:3, nome:'Bolsa não lançada' },
  'sem-mensalidade': { t:'GERAR AS MENSALIDADES', ordem:4, nome:'Contrato sem mensalidade' },
  cancelado: { t:'REFAZER O CONTRATO', ordem:5, nome:'Contrato com a bolsa cancelado' },
  'sem-contrato': { t:'GERAR O CONTRATO', ordem:6, nome:'Matrícula sem contrato' },
  'parcelas-fora': { t:'CONFERIR AS PARCELAS', ordem:7, nome:'Parcelas diferentes de 12' },
  'sem-cota': { t:'LANÇAR A COTA', ordem:8, nome:'Contrato sem cota de entrada' },
  'cancelado-sem-bolsa': { t:'CONFERIR', ordem:9, nome:'Contrato cancelado sem bolsa' },
  'nao-renovou': { t:'CONFERIR', ordem:10, nome:'Proposta entregue, renovação não efetivada' }
};

const lista = [...alunos.values()].filter(a => a.caso).sort((x, y) =>
  CASOS[x.caso].ordem - CASOS[y.caso].ordem ||
  ORDER.indexOf(x.u) - ORDER.indexOf(y.u) ||
  String(x.nome).localeCompare(String(y.nome), 'pt'));

/* Para cada contrato, o que fazer com ele. */
for (const a of lista) {
  const vivos = a.cs.filter(c => !c.fantasma);
  const ativos = vivos.filter(c => c.conta);
  const guardar = ativos.filter(BOM).sort((x, y) => y.pct - x.pct)[0] ||
                  ativos.filter(c => c.bolsa > 0).sort((x, y) => y.pct - x.pct)[0];
  const quieto = c => ({ c, fazer:'—', porque: c.complementar
    ? 'Contraturno (Semi-Integral), contratado à parte da série. Não é trilha em dobro: fica como está.'
    : c.sit === 'CANCELADO'
    ? 'Contrato cancelado' + (c.bolsa > 0 ? ', tinha bolsa de ' + bol(c) + '.' : ', sem bolsa.')
    : 'Sem cobrança própria.' });
  a.linhas = (vivos.length ? vivos : [null]).map(c => {
    if (!c) return { c: null, fazer: CASOS[a.caso].t,
      porque: 'A turma está no cadastro e nenhuma mensalidade foi lançada: o contrato de 2027 nunca foi gerado.' };
    if (a.caso === 'dobro') {
      if (c === guardar) return { c, fazer:'MANTER',
        porque:'É o contrato certo: ' + (c.plano === 'avista' ? 'anuidade à vista' : '12 mensalidades') + ' com a bolsa de ' + bol(c) + '. Fica este.' };
      if (!c.ativo) return quieto(c);
      return { c, fazer:'EXCLUIR O CONTRATO', porque:'Contrato em dobro' +
        (VAZIO(c) ? ', sem nenhuma mensalidade lançada' : ', a preço de tabela cheia') +
        '. O aluno já tem o contrato com bolsa na outra linha: cancelar este.' };
    }
    if (a.caso === 'escolher') return c.ativo
      ? { c, fazer:'ESCOLHER A TRILHA', porque:'O aluno tem contrato ativo nesta trilha, com bolsa de ' + bol(c) +
          ', e em outra. Só vai cursar uma: confirmar com a família e cancelar a que sobrar.' }
      : quieto(c);
    if (a.caso === 'sem-bolsa') return c.ativo && CHEIO(c)
      ? { c, fazer:'LANÇAR A BOLSA', porque:'Contrato ativo a preço de tabela cheia' +
          (a.refPct ? ': lançar a Bolsa Futura de ' + pc(a.refPct) + ' combinada com a família.'
                    : ': nenhuma bolsa foi lançada e não há proposta registrada. Conferir com a unidade.') }
      : c.ativo ? { c, fazer:'EXCLUIR O CONTRATO', porque:'Contrato ativo sem mensalidade, em dobro com a linha de cima: cancelar.' }
      : quieto(c);
    if (a.caso === 'sem-mensalidade') return c.ativo
      ? { c, fazer:'GERAR AS MENSALIDADES', porque:'Contrato ativo' +
          (c.bolsa > 0 ? ' e com a bolsa de ' + bol(c) + ' já lançada' : '') +
          ', mas nenhuma das 12 mensalidades foi gerada: a família não tem boleto nenhum para 2027.' }
      : quieto(c);
    if (a.caso === 'parcelas-fora') return c.ativo && PARCIAL(c)
      ? { c, fazer:'CONFERIR AS PARCELAS', porque:'O contrato tem ' + c.nMens +
          ' mensalidades no lugar de 12. Conferir o plano de pagamento antes de cobrar.' }
      : quieto(c);
    if (a.caso === 'sem-cota') return c.conta && c.plano === 'mensal' && !c.nCotas
      ? { c, fazer:'LANÇAR A COTA', porque:'As 12 mensalidades estão lançadas e nenhuma cota de entrada: a família não tem o boleto da matrícula.' }
      : quieto(c);
    if (a.caso === 'cancelado') return c.bolsa > 0
      ? { c, fazer:'REFAZER O CONTRATO', porque:'Este contrato tinha a bolsa de ' + bol(c) +
          ' e foi cancelado. O aluno ficou sem contrato válido para 2027: refazer com a mesma bolsa.' }
      : quieto(c);
    return { c, fazer: CASOS[a.caso].t, porque: c.bolsa > 0
      ? 'Contrato cancelado com bolsa de ' + bol(c) + '.'
      : 'Contrato cancelado e sem bolsa: conferir se a família desistiu da renovação.' };
  });
  // Quanto do ticket está inflado, na linha do contrato que infla — assim a
  // coluna soma sem contar o mesmo aluno duas vezes.
  const noTicket = a.cs.find(c => c.ativo && c.bolsa <= 0 && c.anual != null);
  a.noTicket = noTicket ? noTicket.anual : null;
  a.anualCerta = noTicket && a.refPct != null ? noTicket.anual * (1 - a.refPct) : null;
  a.inflado = a.anualCerta != null ? noTicket.anual - a.anualCerta : null;
  a.linhas.forEach(o => { if (o.c && o.c === noTicket) {
    o.anualCerta = a.anualCerta; o.inflado = a.inflado; } });
  // Quantas matrículas o aluno tem de pé no cadastro: é o que mostra a dobra.
  a.nAtivas = a.mats.filter(x => AVANCOU.has(x.sit)).length;
}

module.exports = { alunos, lista, CASOS, CODE, ORDER, brl };

/* ---------------- planilha ---------------- */

if (require.main === module) {
  const S = { hdr:1, txt:2, data:3, moeda:4, pct:5, num:6, acao:7, marca:8, titulo:9,
              legenda:10, verde:11, lar:12, larData:13, larMoeda:14, larNum:15 };
  const t = (v, s) => ({ v: v == null ? '' : v, t:'s', s: s == null ? S.txt : s });
  const n = (v, s) => ({ v: v || 0, t:'n', s: s == null ? S.num : s });
  const m = (v, s) => ({ v: Math.round((v || 0) * 100) / 100, t:'m', s: s == null ? S.moeda : s });
  const p = (v, s) => ({ v: Math.round((v || 0) * 10000) / 10000, t:'n', s: s == null ? S.pct : s });

  // Toda linha se sustenta sozinha: unidade, aluno, RA, série e situação
  // repetem em cada contrato, senão o filtro do Excel perde a linha de baixo.
  const CAB = ['Unidade','Aluno','RA','Série','Matrículas ativas','Situação desta matrícula',
    'O QUE FAZER','Por quê','Situação','Trilha do contrato','Contrato','Plano de pagamento',
    'Situação do contrato','Mensalidades','Mensalidade de tabela','Anuidade de tabela',
    'Bolsa lançada','Bolsa %','Anuidade que entra no ticket','Bolsa Futura de referência',
    'Fonte da referência','Anuidade com a bolsa','Infla o ticket em','1ª cobrança','Paga em 2026'];
  const COLS = [18,32,13,8,11,22,20,58,30,30,11,34,14,12,15,15,14,9,16,13,22,16,15,24,13];

  const L = [{ cells: CAB.map(x => t(x, S.hdr)), h:34 }];
  let bloco = 0;
  for (const a of lista) {
    const z = bloco++ % 2 === 1;
    const st = { txt: z ? S.lar : S.txt, num: z ? S.larNum : S.num,
                 moeda: z ? S.larMoeda : S.moeda, pct: z ? S.larNum : S.pct };
    const sits = [...new Set(a.mats.map(x => x.sit))].join(' / ');
    a.linhas.forEach(o => {
      const c = o.c;
      L.push({ cells: [
        t(CODE[a.u] || a.u, st.txt),
        t(a.nome, st.txt),
        t(a.ra, st.txt),
        t(a.serie, st.num),
        n(a.nAtivas, st.num),
        t(c && c.sitMat ? c.sitMat : sits, st.txt),
        t(o.fazer, /EXCLUIR|LANÇAR|REFAZER|GERAR|ESCOLHER/.test(o.fazer) ? S.acao : st.num),
        t(o.porque, st.txt),
        t(CASOS[a.caso].nome, st.txt),
        t(c ? c.trilha : '(sem contrato)', st.txt),
        t(c && c.id ? c.id : '', st.num),
        t(c ? c.planoPg || '' : '', st.txt),
        t(c ? c.sit || '' : '', st.num),
        c ? n(c.nMens, st.num) : t('', st.num),
        c && c.mensal ? m(c.mensal, st.moeda) : t('', st.num),
        c && c.cheia ? m(c.cheia, st.moeda) : t('', st.num),
        c && c.bolsa ? m(c.bolsa, st.moeda) : t('', st.num),
        c && c.bolsa ? p(c.pct, st.pct) : t('', st.num),
        c && c.ativo && c.anual != null ? m(c.anual, st.moeda) : t('', st.num),
        a.refPct == null ? t('', st.num) : p(a.refPct, st.pct),
        t(a.refFonte, st.txt),
        o.anualCerta == null ? t('', st.num) : m(o.anualCerta, st.moeda),
        o.inflado == null ? t('', st.num) : m(o.inflado, st.moeda),
        t(c ? c.cobranca : '', st.txt),
        a.pago26 ? m(a.pago26, st.moeda) : t('', st.num)
      ] });
    });
  }

  /* resumo */
  const L2 = [];
  const tit = x => L2.push({ cells:[t(x, S.titulo)], h:22 });
  const soma = (ls, f) => ls.reduce((s, a) => s + (f(a) || 0), 0);
  const mexer = ls => ls.reduce((s, a) => s + a.linhas.filter(o => o.fazer !== '—' && o.fazer !== 'MANTER').length, 0);
  const CAB2 = ['Situação','Alunos','Contratos a mexer','Anuidade no ticket hoje',
    'Anuidade com a bolsa','Infla o ticket em'];
  const ordem = Object.keys(CASOS).sort((x, y) => CASOS[x].ordem - CASOS[y].ordem);

  tit('Por situação');
  L2.push({ cells: CAB2.map(x => t(x, S.hdr)), h:32 });
  for (const k of ordem) {
    const g = lista.filter(a => a.caso === k); if (!g.length) continue;
    L2.push({ cells: [t(CASOS[k].nome), n(g.length), n(mexer(g)),
      m(soma(g, a => a.inflado != null ? a.noTicket : 0)), m(soma(g, a => a.anualCerta)), m(soma(g, a => a.inflado))] });
  }
  L2.push({ cells: [t('Rede', S.verde), n(lista.length, S.verde), n(mexer(lista), S.verde),
    m(soma(lista, a => a.inflado != null ? a.noTicket : 0), S.verde), m(soma(lista, a => a.anualCerta), S.verde),
    m(soma(lista, a => a.inflado), S.verde)] });

  L2.push({ cells:[] });
  tit('Por unidade');
  L2.push({ cells: ['Unidade','Alunos','Contratos a mexer','Anuidade no ticket hoje',
    'Anuidade com a bolsa','Infla o ticket em'].map(x => t(x, S.hdr)), h:32 });
  for (const u of ORDER) {
    const g = lista.filter(a => a.u === u); if (!g.length) continue;
    L2.push({ cells: [t(CODE[u]), n(g.length), n(mexer(g)),
      m(soma(g, a => a.inflado != null ? a.noTicket : 0)), m(soma(g, a => a.anualCerta)), m(soma(g, a => a.inflado))] });
  }

  const L3 = [];
  [['Como ler', S.titulo],
   ['A Bolsa Futura é a proposta de renovação entregue à família. Esta planilha reúne os alunos de rematrícula de 2027 cujo contrato ficou a preço de tabela cheia, ou que ficaram sem contrato válido.', S.legenda],
   ['Cada aluno aparece com uma linha por contrato, na mesma faixa de cor. Unidade, aluno, RA, série e situação repetem em todas as linhas dele, para que o filtro do Excel não perca nenhuma.', S.legenda],
   ['"Situação desta matrícula" é o status da trilha daquela linha, não do aluno: é por isso que o mesmo aluno pode aparecer Matriculado numa trilha e Pré-Matriculado na outra. "Matrículas ativas" conta quantas ele tem de pé.', S.legenda],
   ['"Anuidade com a bolsa" e "Infla o ticket em" aparecem só na linha do contrato que infla o ticket, para a coluna somar sem contar o mesmo aluno duas vezes.', S.legenda],
   ['O QUE FAZER', S.titulo],
   ['EXCLUIR O CONTRATO — contrato em dobro. O aluno escolheu uma trilha, depois trocou, e os dois contratos ficaram ativos. Fica o que tem a bolsa; cancelar o de tabela cheia.', S.legenda],
   ['ESCOLHER A TRILHA — o aluno tem contrato ativo em duas trilhas e as duas têm bolsa. Não há linha errada para cancelar: a família escolhe, e a outra sai.', S.legenda],
   ['LANÇAR A BOLSA — contrato ativo, correto no mais, mas sem nenhuma bolsa lançada. A família está com uma anuidade de tabela cheia no boleto.', S.legenda],
   ['REFAZER O CONTRATO — o contrato que tinha a bolsa foi cancelado e nada entrou no lugar. O aluno ficou sem contrato de 2027.', S.legenda],
   ['GERAR O CONTRATO — a matrícula está fechada no cadastro e nenhuma mensalidade foi lançada.', S.legenda],
   ['LANÇAR A COTA — as 12 mensalidades estão lançadas e falta a cota de entrada: a família não tem o boleto da matrícula.', S.legenda],
   ['Linha com — não pede ação. A mais comum é o contraturno (Semi-Integral / CCT), que é serviço contratado à parte da série: dois contratos assim não são trilha em dobro.', S.legenda],
   ['CONFERIR — contrato cancelado sem bolsa, ou proposta entregue e renovação não efetivada: antes de mexer, confirmar com a família.', S.legenda],
   ['As contas', S.titulo],
   ['Anuidade de tabela é a soma do VALOR_ORIGINAL das 12 mensalidades mais a cota de entrada. Bolsa lançada é a soma de BOLSA/DEDUÇÕES. O que a família paga é a diferença — a mesma conta que o painel usa no ticket.', S.legenda],
   ['"Anuidade que entra no ticket" é o que este contrato leva hoje para o ticket médio da renovação: 12 mensalidades líquidas mais R$ 300 de cota (R$ 600 no Américas).', S.legenda],
   ['As colunas de dinheiro contam só os alunos que têm bolsa de referência: nos outros não há com o que comparar, e eles entram apenas na contagem.', S.legenda],
   ['"Infla o ticket em" é a diferença entre a anuidade de hoje e a anuidade com a bolsa de referência. Enquanto o contrato errado está de pé, esse valor empurra o ticket médio da rede para cima. Corrigir faz o ticket cair: é o número verdadeiro aparecendo.', S.legenda],
   ['A Bolsa Futura de referência vem da lista do processo de renovação quando o aluno está nela; nos outros casos vem da bolsa que o próprio aluno tem em outro contrato. A coluna "Fonte da referência" diz qual das duas.', S.legenda],
   ['"Paga em 2026" é o que a família paga neste ano, da lista da Bolsa Futura. Serve para a conversa com o responsável; não sai do Totvs.', S.legenda],
   ['Linhas sem lançamento financeiro (VALOR_ORIGINAL vazio e boleto SEM BOLETO) são resíduo do export e não contam como contrato.', S.legenda],
   ['Ficha cadastral e ficha financeira de ' + DIA + ' de 2026.', S.legenda]
  ].forEach(([x, s]) => L3.push({ cells:[t(x, s)], h: s === S.titulo ? 22 : undefined }));

  const buf = pasta([
    { nome:'Para resolver', cols:COLS, linhas:L, congelar:1,
      filtro: 'A1:' + colName(COLS.length - 1) + L.length },
    { nome:'Resumo', cols:[34,10,18,22,20,18], linhas:L2 },
    { nome:'Como ler', cols:[130], linhas:L3 }
  ]);
  const destino = path.join(SAIDA, 'Renovação 2027 - alunos sem a Bolsa Futura.xlsx');
  fs.writeFileSync(destino, buf);

  console.log('->', destino, Math.round(buf.length / 1024) + ' KB');
  console.log(lista.length + ' alunos · ' + (L.length - 1) + ' linhas de contrato');
  const ord = Object.keys(CASOS).sort((x, y) => CASOS[x].ordem - CASOS[y].ordem);
  for (const k of ord) {
    const g = lista.filter(a => a.caso === k); if (!g.length) continue;
    console.log('  ' + CASOS[k].nome.padEnd(44) + String(g.length).padStart(3) + ' alunos · infla ' +
      brl(soma(g, a => a.inflado)));
  }
  console.log('  ' + 'TOTAL'.padEnd(44) + String(lista.length).padStart(3) + ' alunos · infla ' +
    brl(soma(lista, a => a.inflado)));
  const fora = Object.keys(PROPOSTA).filter(ra => !lista.some(a => a.ra === ra));
  if (fora.length) console.log('\n  da lista da Bolsa Futura, fora da planilha: ' + fora.join(' '));
}
