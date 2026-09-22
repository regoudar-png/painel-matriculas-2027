// Cruza as três bases e devolve o snapshot que alimenta o painel.
//
//   node motor/cruzar.js <totvs.xlsx> <planilha-unidades.xlsx> [marketplace.xlsx] > snapshot.json
//
// O Marketplace é opcional: sem ele o cruzamento roda só com Totvs + planilha,
// e todo aluno fica sem informação de pagamento (o painel avisa).

const fs = require('fs');
const path = require('path');
const { lerPlanilha } = require('./ler-xlsx.js');

// Caminho de arquivo (linha de comando) ou planilha já lida (navegador, que
// descompacta o .xlsx antes de chamar o motor). Aceita os dois.
function abrir(x){ return (x && x.abas) ? x : lerPlanilha(x); }

/* ---------------- unidades ---------------- */

const CODE = {
  BG:'Bangu', CG:'Campo Grande', CX:'Duque de Caxias', MD:'Madureira', NI:'Nova Iguaçu',
  RM:'Rocha Miranda', RT:'Retiro dos Artistas', SJ:'São João de Meriti', TQ:'Taquara',
  TJ:'Tijuca', AM:'Américas'
};
const ORDER = ['BG','CG','AM','CX','RM','TQ','TJ','NI','MD','RT','SJ'];
// Américas aparece no Totvs como "Recreio".
const TOTVS_PARA_CODIGO = {
  'TAQUARA':'TQ','DUQUE DE CAXIAS':'CX','BANGU':'BG','ROCHA MIRANDA':'RM','NOVA IGUACU':'NI',
  'SÃO JOÃO DE MERITI':'SJ','CAMPO GRANDE':'CG','RETIRO DOS ARTISTAS':'RT','TIJUCA':'TJ',
  'MADUREIRA':'MD','RECREIO':'AM'
};
const MERCADO_PARA_CODIGO = {};
for (const k in CODE) MERCADO_PARA_CODIGO[CODE[k]] = k;
// A Layers escreve a unidade de dois jeitos: "Campo Grande" no export antigo e
// "Bolsão 2027 - Campo Grande" no novo. E abrevia Duque de Caxias.
MERCADO_PARA_CODIGO['Caxias'] = 'CX';
const unidadeMkt = v => {
  const s = String(v || '').replace(/^Bols[ãa]o\s*\d*\s*-\s*/i, '').trim();
  return MERCADO_PARA_CODIGO[s] || '??';
};

/* ---------------- situações do Totvs ---------------- */

// Regra da coordenação: só conta como MATRÍCULA quem está "Matriculado".
const MATRICULADO = new Set(['Matriculado','Reserva de vaga - Matriculado']);
// Contrato assinado, matrícula ainda não fechada.
const PRE = new Set(['Pré-Matriculado','Reserva de vaga - Pré-Mat']);
const PENDENTE = new Set(['Pendente','Reserva de vaga - Pendente']);
const ENCERRADO = new Set(['Cancelado','Reserva de vaga - Cancelada','Desistente']);
const avancou = s => MATRICULADO.has(s) || PRE.has(s);

/* ---------------- comparação de nomes ---------------- */

const norm = s => (s||'').toString().normalize('NFD').replace(/[̀-ͯ]/g,'')
  .toUpperCase().replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();
const SEM_PESO = new Set(['DE','DA','DO','DAS','DOS','E']);
const pedacos = s => norm(s).split(' ').filter(t => t.length>1 && !SEM_PESO.has(t));

function bigramas(s){ const o=new Set(); for(let i=0;i<s.length-1;i++) o.add(s.slice(i,i+2)); return o; }

// Os bigramas de cada palavra são sempre os mesmos; guardo em vez de refazer.
const _big = new Map();
const bigCache = s => { let b = _big.get(s); if(!b){ b = bigramas(s); _big.set(s, b); } return b; };
function dice(a,b){
  const A=bigCache(a), B=bigCache(b);
  if(!A.size||!B.size) return 0;
  let i=0; for(const g of A) if(B.has(g)) i++;
  return 2*i/(A.size+B.size);
}

// Cada nome é normalizado, quebrado e ordenado uma vez só. Sem isso o mesmo
// nome era reprocessado a cada uma das centenas de comparações em que entra.
const _prep = new Map();
function preparar(x){
  const chave = x == null ? '' : String(x);
  let p = _prep.get(chave);
  if(p) return p;
  const n = norm(chave);
  const t = pedacos(chave);
  p = { n, t, ordenado: t.slice().sort().join(' '), colado: n.replace(/ /g,'') };
  _prep.set(chave, p);
  return p;
}

function parecenca(a,b){
  const pa=preparar(a), pb=preparar(b);
  const na=pa.n, nb=pb.n;
  if(!na||!nb) return 0;
  if(na===nb) return 1;
  const ta=pa.t, tb=pb.t;
  if(!ta.length||!tb.length) return 0;
  if(pa.ordenado===pb.ordenado) return 0.98;
  let acerto=0; const usados=new Set();
  for(const t of ta){
    let melhor=-1, nota=0;
    tb.forEach((u,j)=>{ if(usados.has(j))return; const s = t===u?1:dice(t,u); if(s>nota){nota=s;melhor=j;} });
    if(nota>=0.75){ acerto+=nota; usados.add(melhor); }
  }
  const cobertura = acerto/Math.min(ta.length,tb.length);
  const d = dice(pa.colado, pb.colado);
  const primeiroBate = ta[0]&&tb[0]&&(ta[0]===tb[0]||dice(ta[0],tb[0])>=0.7);
  let s = Math.max(cobertura*0.55 + d*0.45, d);
  if(!primeiroBate) s -= 0.10;
  return s;
}
const LIMITE = 0.80;   // une sozinho
const SUGERE = 0.58;   // só sugere
// Na faixa de dúvida, o responsável precisa bater bem para o par valer.
const CONFIRMA_RESP = 0.75;

/* ---------------- leitura das bases ---------------- */

const serial = n => {
  const v = parseFloat(n);
  if(!isFinite(v)) return null;
  return new Date(Date.UTC(1899,11,30) + Math.floor(v)*86400000).toISOString().slice(0,10);
};
const isoDe = s => {
  if(!s) return null;
  const m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? m[3]+'-'+m[2]+'-'+m[1] : null;
};
// A planilha mistura formatos entre abas: "300,00" numa, "300.0" noutra.
// É preciso descobrir o separador decimal em vez de presumir — presumir
// vírgula transformava 660.0 em 6600.
const num = v => {
  if (v == null || v === '') return null;
  let s = String(v).trim().replace(/[R$\s]/g, '');
  if (!s || !/[0-9]/.test(s)) return null;
  const temPonto = s.includes('.'), temVirgula = s.includes(',');
  if (temPonto && temVirgula) {
    // o separador que aparece por último é o decimal
    s = s.lastIndexOf(',') > s.lastIndexOf('.')
      ? s.replace(/\./g, '').replace(',', '.')
      : s.replace(/,/g, '');
  } else if (temVirgula) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (temPonto) {
    // ponto sozinho: 3 dígitos depois dele e um único ponto = milhar
    const p = s.split('.');
    if (p.length === 2 && p[1].length === 3 && p[0].length <= 3) s = p.join('');
    else if (p.length > 2) s = p.join('');
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
};

function lerTotvs(caminho){
  const { abas } = abrir(caminho);
  const aba = abas.filter(a=>a.linhas).sort((a,b)=>b.linhas.length-a.linhas.length)[0];
  const dados = aba.linhas.slice(1).filter(r => r && r[13]);
  const relato = { linhas: dados.length };

  // 1) captação é só TIPO = MATRÍCULA. Rematrícula é outra operação e vai
  //    inteira para a renovação — inclusive a que já avançou de situação.
  //    O Semi-Integral (curso CCT, séries SINTA/SINTB) é contratado por cima
  //    da matrícula regular e o Totvs registra como uma segunda MATRÍCULA. Não
  //    é captação: o aluno da casa que o contrata aparecia como pré-matrícula
  //    sem pagamento — foi o que o Retiro apontou.
  const complementar = r => String(r[4] || '').trim() === 'CCT' || /semi-?integral/i.test(String(r[7] || ''));
  //    E o aluno da casa: o Totvs gera uma linha de REMATRÍCULA para todo
  //    aluno de 2026 (usuário IMPORTADOR). Quando a unidade abre uma MATRÍCULA
  //    nova para um desses, é erro de lançamento — a Mariana Lima Machado
  //    apareceu na captação assim. Quem tem REMATRÍCULA não é captação.
  const daCasa = new Set(dados.filter(r => r[10] === 'REMATRÍCULA' && r[12]).map(r => r[12]));
  const semLiberadas = dados.filter(r => r[10] === 'MATRÍCULA' && !complementar(r) && !daCasa.has(r[12]));
  relato.rematricula = dados.filter(r => r[10] !== 'MATRÍCULA').length;
  relato.complementares = dados.filter(r => r[10] === 'MATRÍCULA' && complementar(r)).length;
  // Quem foi tirado por esta regra fica no relato, para a unidade corrigir o tipo.
  relato.casaComoMatricula = dados.filter(r => r[10] === 'MATRÍCULA' && !complementar(r) && daCasa.has(r[12]))
    .map(r => ({ aluno: String(r[13]).trim(), ra: r[12], sit: r[9], serie: r[5],
      unidade: String(r[1]).replace(/COL[ÉE]GIO E CURSO MATRIZ EDUCA[ÇC][ÃA]O( -)? /,'').trim() }));

  // 2) uma linha por aluno, ficando com a situação mais avançada
  const posto = s => MATRICULADO.has(s) ? 4 : PRE.has(s) ? 3 : PENDENTE.has(s) ? 2 : 1;
  const porRA = new Map();
  for(const r of semLiberadas){
    const k = r[12] || ('N:'+norm(r[13]));
    const atual = porRA.get(k);
    if(!atual || posto(r[9]) > posto(atual[9])) porRA.set(k, r);
  }
  relato.linhasRepetidas = semLiberadas.length - porRA.size;

  // 3) fora os cadastros de teste das unidades
  const todos = [...porRA.values()];
  const reais = todos.filter(r => !/\bteste\b/i.test(r[13]));
  relato.cadastrosTeste = todos.length - reais.length;
  relato.alunos = reais.length;

  const unidade = r => TOTVS_PARA_CODIGO[
    String(r[1]).replace(/COL[ÉE]GIO E CURSO MATRIZ EDUCA[ÇC][ÃA]O( -)? /,'').trim()] || '??';

  const captacao = reais.map(r => ({
    code: unidade(r), aluno: String(r[13]).trim(), ra: r[12],
    serie: r[5], turma: r[6], grade: r[7]||null, turno: r[8], curso: r[4],
    sit: r[9], tipo: r[10], dtMat: serial(r[11]),
    respFin: r[30]||null, cpfResp: r[37]||null
  }));

  // renovação: o universo elegível são as rematrículas, incluindo as só liberadas
  const remat = dados.filter(r => r[10]==='REMATRÍCULA' && !/\bteste\b/i.test(r[13]));
  const porAluno = new Map();
  for(const r of remat){
    const k = r[12] || ('N:'+norm(r[13]));
    const atual = porAluno.get(k);
    if(!atual || posto(r[9]) > posto(atual[9])) porAluno.set(k, r);
  }
  const renovacao = [...porAluno.values()].map(r => ({
    code: unidade(r), aluno: String(r[13]).trim(), ra: r[12],
    serie: r[5], curso: r[4], sit: r[9],
    matriculado: MATRICULADO.has(r[9]), pre: PRE.has(r[9])
  }));

  return { captacao, renovacao, relato };
}

function lerUnidades(caminho){
  // Sem a planilha o painel roda só com Totvs, Layers e financeiro; as listas
  // que dependem dela ficam vazias e o ticket da captação some.
  if(!caminho) return [];
  const { abas } = abrir(caminho);
  const saida = [];
  for(const a of abas){
    const code = a.nome.trim();
    if(code==='Resumo' || !a.linhas || !CODE[code]) continue;
    a.linhas.slice(1).forEach((r,i) => {
      const nome = (r[0]||'').toString().trim();
      if(!nome || nome==='85' || /^\d+([.,]\d+)?$/.test(nome)) return;
      let fluxo = (r[2]||'').toString().trim();
      if(fluxo==='85' || !fluxo) fluxo = '(em branco)';
      else if(/portal/i.test(fluxo)) fluxo = 'Portal do aluno';
      else if(/market/i.test(fluxo)) fluxo = 'Marketplace';
      let bolsa = num(r[5]); if(bolsa!=null && bolsa>1.5) bolsa /= 100;
      const cota = num(r[3]), mensalidade = num(r[4]);
      // Ticket = anuidade: mensalidade (já com a bolsa) × 12 mais a 1ª cota.
      // Bate com a coluna Anuidade da planilha — conferido em 191 de 192 linhas.
      const anuidade = num(r[6]);
      const ticket = mensalidade != null ? mensalidade * 12 + (cota || 0) : anuidade;
      saida.push({ code, linha: i+2, aluno: nome.replace(/[.\s]+$/,''),
        turma: (r[1]||'').toString().trim(), fluxo, cota, mensalidade, bolsa, ticket, anuidade,
        obs: (r[7]||'').toString().trim() || null });
    });
  }
  return saida;
}

const POSTO_PGTO = {'Pago':4,'Pendente / Vencido':3,'Não classificado':2,'Estornado':1};

// O export antigo trazia uma coluna Status já normalizada. O novo traz só o
// status cru da Layers, então refaço a mesma normalização aqui — conferida
// contra o arquivo antigo, onde as duas colunas convivem.
const STATUS_LAYERS = {
  'Pago':'Pago', 'Recebido':'Pago',
  'Vencido':'Pendente / Vencido', 'Em aberto':'Pendente / Vencido',
  'Reembolsado':'Estornado',
  'Falhou':'Não classificado'
};
const dataSerial = v => {
  const x = parseFloat(String(v == null ? '' : v).replace(',', '.'));
  if(!isFinite(x) || x < 20000) return String(v || '').trim();
  const d = new Date(Date.UTC(1899, 11, 30) + x * 864e5);
  return String(d.getUTCDate()).padStart(2,'0') + '/' +
    String(d.getUTCMonth()+1).padStart(2,'0') + '/' + d.getUTCFullYear();
};

function lerMarketplace(caminho){
  if(!caminho) return { pessoas: [], relato: { pedidos: 0, ausente: true } };
  const { abas } = abrir(caminho);
  const aba = abas.filter(a=>a.linhas&&a.linhas.length>2)[0];
  const cab = (aba.linhas[0]||[]).map(v => String(v==null?'':v).trim());
  // "Itens da venda" é o export cru da Layers; o antigo já vinha tratado.
  const novo = cab.indexOf('Código da Venda') >= 0 && cab.indexOf('Status da Venda') >= 0;
  const brutos = [];
  aba.linhas.slice(1).forEach(r => {
    if(novo){
      const aluno = String(r[20]||'').trim();
      if(!aluno) return;
      const bruto = String(r[26]||'').trim();
      brutos.push({ unidade:r[30], code: unidadeMkt(r[30]), aluno,
        seriePret:r[24], resp:String(r[22]||r[3]||'').trim(), tel:r[23], email:r[4], cpf:r[6],
        produto:r[8], valor:+r[1]||0,
        status: STATUS_LAYERS[bruto] || 'Não classificado', statusOrig: bruto,
        metodo:r[28], parcelas:r[27], data: dataSerial(r[2]),
        pedido:r[37], transacao:r[0] });
      return;
    }
    if(!r[1] || r[0]==='Não há venda') return;
    brutos.push({ unidade:r[0], code: unidadeMkt(r[0]), aluno:(r[1]||'').trim(),
      seriePret:r[4], resp:r[5], tel:r[6], email:r[7], cpf:r[8], produto:r[10],
      valor:+r[11]||0, status:r[12], statusOrig:r[13], metodo:r[14], parcelas:r[15],
      data:r[16], pedido:r[18], transacao:r[20] });
  });
  // O export repete algumas transações: mesma linha, mesmo ID. Guardo as duas
  // contagens — a de linhas, que é o que a Layers mostra na tela, e a de
  // transações distintas, que é o dinheiro de verdade.
  // Irmãos entram na mesma venda, com o mesmo Código da Venda: a chave da
  // repetição precisa levar o aluno junto, senão o segundo irmão some do
  // pagamento — foi assim que o Retiro apareceu com matriculado "sem pagamento".
  const vistos = new Set();
  const repetidas = [];
  const unicos = brutos.filter(o => {
    const k = (o.transacao || o.pedido) + '|' + norm(o.aluno);
    if(vistos.has(k)){ repetidas.push(o); return false; }
    vistos.add(k); return true;
  });
  // agrupa por aluno, ficando com o melhor status
  const mapa = new Map();
  for(const o of unicos){
    const k = o.code+'|'+norm(o.aluno);
    if(!mapa.has(k)) mapa.set(k, { ...o, pedidos: [] });
    mapa.get(k).pedidos.push(o);
  }
  const pessoas = [...mapa.values()].map(p => {
    p.pedidos.sort((a,b)=>(POSTO_PGTO[b.status]||0)-(POSTO_PGTO[a.status]||0));
    const melhor = p.pedidos[0];
    return { ...p, status: melhor.status, statusOrig: melhor.statusOrig, valor: melhor.valor,
      metodo: melhor.metodo, parcelas: melhor.parcelas, data: melhor.data, pedido: melhor.pedido,
      nPedidos: p.pedidos.length };
  });
  return { pessoas, repetidas, relato: {
    pedidos: brutos.length, unicos: unicos.length, alunos: pessoas.length,
    linhasPagas: brutos.filter(o=>o.status==='Pago').length,
    transacoesPagas: unicos.filter(o=>o.status==='Pago').length,
    repetidas: repetidas.length,
    repetidasPagas: repetidas.filter(o=>o.status==='Pago').length
  } };
}


/* ---------------- financeiro do Totvs (Portal do aluno) ---------------- */

// A 1ª cota aparece como serviço "1ª Cota de Mensalidade" ou "Reserva de vaga".
// Está paga quando tem DATA_BAIXA ou o boleto consta como baixado.
const SERVICO_ENTRADA = /1ª Cota de Mensalidade|Reserva de vaga/i;
// Mensalidade pura — não a 1ª cota, não a reserva. Só quem tem contrato
// lançado no Totvs aparece aqui; quem está apenas liberado não tem parcela.
const SERVICO_MENSALIDADE = /(^|- )Mensalidade$/i;

// Valor da 1ª cota: R$ 300 nas unidades, R$ 600 em Américas (Recreio).
// A mensalidade não entra nesta conta — o que se rastreia é a cota de entrada.
const COTA_PADRAO = 300, COTA_RECREIO = 600;
const cotaEsperada = filial => /RECREIO/i.test(String(filial||'')) ? COTA_RECREIO : COTA_PADRAO;

function lerFinanceiro(caminho){
  // Sem a ficha o painel roda sem os rastreios financeiros — mapas vazios e a
  // marca de ausência, para ninguém aparecer como 'sem financeiro' por engano.
  if(!caminho) return { porRA: new Map(), mensPorRA: new Map(), cobranca: new Map(), relato: { ausente: true } };
  const { abas } = abrir(caminho);
  const aba = abas.filter(a=>a.linhas).sort((a,b)=>b.linhas.length-a.linhas.length)[0];
  const linhas = aba.linhas.slice(1).filter(r => r && r[3]);
  const entrada = linhas.filter(r => SERVICO_ENTRADA.test(String(r[34]||'')));
  // DATA_BAIXA traz lixo numérico ("157", "2504") na maior parte das linhas.
  const ehData = v => /^\d{2}\/\d{2}\/\d{4}$/.test(String(v||'').trim());

  // Um aluno pode ter vários lançamentos de entrada; o que vale é a soma do
  // que foi efetivamente baixado.
  const grupos = new Map();
  for(const r of entrada){
    if(!grupos.has(r[3])) grupos.set(r[3], []);
    grupos.get(r[3]).push(r);
  }

  const porRA = new Map();
  for(const [ra, rs] of grupos){
    // Uma parcela pode vir duas vezes no export (mesmo REF_FINANCEIRO, um
    // status de boleto em cada linha). Somar as duas dobraria o valor pago.
    const vistos = new Set();
    const baixas = rs.filter(r => {
      if(!((num(r[44])||0) > 0 && ehData(r[45]))) return false;
      const ref = String(r[35]||'').trim();
      if(!ref) return true;
      if(vistos.has(ref)) return false;
      vistos.add(ref); return true;
    });
    const total = baixas.reduce((a,r) => a + (num(r[44])||0), 0);
    const esperado = cotaEsperada(rs[0][2]);
    const ultima = baixas.length ? baixas[baixas.length-1] : null;
    porRA.set(ra, {
      ra, aluno: String(rs[0][4]||'').trim(), tipo: rs[0][19],
      servico: (ultima||rs[0])[34], esperado, valor: total,
      dtBaixa: ultima ? String(ultima[45]).trim() : null,
      venc: String(rs[0][41]||'').trim() || null,
      statusBoleto: (ultima||rs[0])[47] || null,
      // A descrição traz a trilha que falta na base de matrículas.
      descricao: rs[0][17] || null, turma: rs[0][23] || null,
      pago: total >= esperado * 0.98,
      parcial: total > 0 && total < esperado * 0.98 ? { esperado, pago: total, falta: esperado - total } : null
    });
  }

  // Retrato da cobrança de cada aluno: quantos lançamentos de entrada existem,
  // se o plano do ano foi gerado, e o que está duplicado. É daqui que saem os
  // rastreios de erro de lançamento.
  const cobranca = new Map();
  const vistosRef = new Set();
  // Linha sem lançamento financeiro: o export traz VALOR_ORIGINAL vazio,
  // REF_FINANCEIRO 4837 e VALOR_BAIXADO 4838 em todas — número de controle do
  // relatório, não cobrança. Contar isso como entrada inflava a dobra.
  const fantasma = r => (num(r[42]) || 0) < 0.005 && String(r[47] || '') === 'SEM BOLETO';
  for(const r of linhas){
    const serv = String(r[34] || '');
    const ehEntrada = SERVICO_ENTRADA.test(serv);
    const ehMens = SERVICO_MENSALIDADE.test(serv);
    if(!ehEntrada && !ehMens) continue;
    if(!cobranca.has(r[3])) cobranca.set(r[3], { entradas: [], mensalidades: 0 });
    const o = cobranca.get(r[3]);
    const status = String(r[47] || '');
    // Cancelado no Totvs já foi resolvido pela unidade; não é cobrança viva.
    if(status === 'Cancelado' || fantasma(r)) continue;
    if(ehMens){ o.mensalidades++; continue; }
    // A mesma parcela vem duas vezes quando o boleto tem dois status de
    // remessa; isso é repetição do export, não cobrança em dobro.
    const ref = String(r[35] || '').trim();
    if(ref){ if(vistosRef.has(ref)) continue; vistosRef.add(ref); }
    const pago = num(r[44]) || 0;
    o.entradas.push({ servico: serv, venc: String(r[41] || '').trim(),
      tipo: /Reserva de vaga/i.test(serv) ? 'reserva' : 'cota',
      valor: (num(r[42]) || 0) - (num(r[43]) || 0), ref, plano: r[25] || null,
      baixado: pago > 0 && ehData(r[45]),
      // Bolsa de 100% aplicada e baixa a zero: a unidade anulou a cobrança.
      zerada: status === 'Baixado' && pago < 0.005 });
  }

  // Mensalidade líquida por aluno: valor original menos as deduções da bolsa.
  // Uso a mediana das parcelas porque a primeira e a última costumam variar.
  const mensPorRA = new Map();
  const porAlunoMens = new Map();
  for(const r of linhas){
    if(!SERVICO_MENSALIDADE.test(String(r[34]||''))) continue;
    const liq = (num(r[42])||0) - (num(r[43])||0);
    if(liq <= 0) continue;
    if(!porAlunoMens.has(r[3])) porAlunoMens.set(r[3], { filial:r[2], tipo:r[19], vals:[], ded:0, cad:'' });
    const a0 = porAlunoMens.get(r[3]);
    a0.vals.push(liq);
    a0.ded += num(r[43]) || 0;
    if(!a0.cad && String(r[48]||'').trim()) a0.cad = String(r[48]).trim();
  }
  for(const [ra, a] of porAlunoMens){
    a.vals.sort((x,y)=>x-y);
    const mediana = a.vals[Math.floor(a.vals.length/2)];
    const soma = a.vals.reduce((x,y)=>x+y, 0);
    // Mensalidade do ano: a soma das parcelas lançadas. Quando o contrato não
    // tem as 12, normalizo pela mediana × 12 — senão um contrato pela metade
    // viraria meia anuidade. A cota entra no join, com a unidade da matrícula.
    const anoMens = a.vals.length === 12 ? soma : mediana * 12;
    // Sem dedução nenhuma e sem bolsa cadastrada: a anuidade está bruta, a
    // bolsa não foi lançada. Fica marcado para a unidade corrigir na origem.
    mensPorRA.set(ra, { mensal: mediana, soma, nParc: a.vals.length, anoMens,
      ded: a.ded, bolsaCad: a.cad || null, semBolsa: a.ded === 0 && !a.cad,
      tipo: a.tipo, filial: a.filial });
  }

  const todos = [...porRA.values()];
  return { porRA, mensPorRA, cobranca, relato: {
    linhas: linhas.length, entradas: entrada.length, alunos: todos.length,
    cotaCompleta: todos.filter(x=>x.pago).length,
    parciais: todos.filter(x=>x.parcial).length,
    semBaixa: todos.filter(x=>!x.valor).length,
    comMensalidade: mensPorRA.size
  }};
}

/* ---------------- pareamento ---------------- */

// Nome do responsável de cada lado, quando existe. É o desempate quando o nome
// do aluno vem escrito diferente nas duas bases.
const respDe = o => o && (o.resp || o.respFin || o.responsavel) || '';

// Une duas listas por nome dentro da mesma unidade, do par mais parecido para o
// menos. Na faixa de dúvida (SUGERE a LIMITE) o par só entra se o responsável
// confirmar — sem isso a Amanda do Américas ficava de fora e a Julia da Taquara
// entraria errada.
// Palavras do nome que servem de chave de busca. Fora as partículas, que
// aparecem em quase todo nome e não separam ninguém.
const PARTICULA = new Set(['de','da','do','das','dos','e','di','du','del','la','van','von']);
function chaves(nome){
  return norm(nome).split(' ').filter(p => p.length >= 3 && !PARTICULA.has(p));
}

function parear(a, b, nomeA, nomeB){
  // Índice por palavra: só comparo quem divide alguma palavra do nome. Dois
  // nomes com semelhança suficiente para virar par sempre dividem uma.
  const balde = new Map();
  b.forEach((y, j) => chaves(y[nomeB]).forEach(k => {
    if(!balde.has(k)) balde.set(k, []);
    balde.get(k).push(j);
  }));

  const pares = [];
  a.forEach((x, i) => {
    const vistos = new Set();
    chaves(x[nomeA]).forEach(k => (balde.get(k) || []).forEach(j => vistos.add(j)));
    vistos.forEach(j => {
      const y = b[j];
      const s = parecenca(x[nomeA], y[nomeB]);
      if(s >= LIMITE){ pares.push({ i, j, s, mesma: x.code===y.code }); return; }
      if(s < SUGERE) return;
      const ra = respDe(x), rb = respDe(y);
      if(!ra || !rb) return;
      const r = parecenca(ra, rb);
      if(r >= CONFIRMA_RESP) pares.push({ i, j, s, r, mesma: x.code===y.code, viaResp: true });
    });
  });
  pares.sort((p,q) => (q.mesma-p.mesma) || (q.s-p.s));
  const usadoA = new Set(), usadoB = new Set();
  for(const p of pares){
    if(usadoA.has(p.i) || usadoB.has(p.j)) continue;
    usadoA.add(p.i); usadoB.add(p.j);
    a[p.i]._par = b[p.j]; b[p.j]._par = a[p.i];
    a[p.i]._nota = b[p.j]._nota = Math.round(p.s*100);
    if(p.viaResp){ a[p.i]._viaResp = b[p.j]._viaResp = Math.round(p.r*100); }
    if(!p.mesma){ a[p.i]._outraUnidade = true; b[p.j]._outraUnidade = true; }
  }
}

// Mesma pessoa aparecendo duas vezes no Marketplace: mesmo aluno, mesma
// unidade, nome quase igual. Fico com o registro de melhor status e guardo os
// outros para o painel não perder o histórico do estorno.
const ORDEM_STATUS = { 'Pago':4, 'Pendente / Vencido':3, 'Estornado':2 };
function juntarRepetidos(pessoas){
  const forca = p => ORDEM_STATUS[p.status] || 1;
  for(let i = 0; i < pessoas.length; i++){
    const x = pessoas[i];
    if(x._absorvido) continue;
    for(let j = i + 1; j < pessoas.length; j++){
      const y = pessoas[j];
      if(y._absorvido || y.code !== x.code) continue;
      if(parecenca(x.aluno, y.aluno) < 0.90) continue;
      const [fica, sai] = forca(x) >= forca(y) ? [x, y] : [y, x];
      sai._absorvido = true;
      (fica._irmaos = fica._irmaos || []).push({ status: sai.status, pedido: sai.pedido,
        valor: sai.valor, data: sai.data });
    }
  }
  for(let i = pessoas.length - 1; i >= 0; i--) if(pessoas[i]._absorvido) pessoas.splice(i, 1);
}

/* ---------------- categorias ---------------- */

const CATEGORIAS = [
  { id:'falta_totvs', chip:'FALTA LANÇAR', k:'crit',
    t:'Pagou a cota e não está no Totvs',
    d:'A 1ª cota foi paga e confirmada, mas não existe matrícula 2027 para esse aluno no Totvs. É o que trava a matrícula — a unidade precisa lançar.' },
  { id:'totvs_pendente', chip:'CONCLUIR', k:'warn',
    t:'No Totvs, matrícula ainda pendente',
    d:'O aluno já existe no Totvs para 2027, mas a situação é Pendente — o lançamento foi aberto e não foi concluído.' },
  { id:'totvs_cancel', chip:'CANCELADO', k:'warn',
    t:'No Totvs como cancelado ou desistente',
    d:'Há registro no Totvs, mas cancelado ou com desistência. Se a cota foi paga, decidir entre estorno e retomada.' },
  { id:'falta_pagamento', chip:'SEM PAGAMENTO', k:'warn',
    t:'Na planilha da unidade, sem pagamento nem Totvs',
    d:'A unidade lançou o aluno na planilha, mas não há pagamento no Marketplace nem no Portal do aluno, e não há matrícula no Totvs. Quem está marcado FLUXO EM BRANCO tem também a coluna FLUXO vazia na planilha — o registro parou antes de qualquer coisa acontecer.' },
  { id:'pgto_pendente', chip:'COBRANÇA', k:'warn',
    t:'Pagamento não concluído',
    d:'Cota gerada e ainda vencida ou em aberto. Cobrança pendente com a família.' },
  { id:'estorno', chip:'ESTORNO', k:'warn',
    t:'Pagamento estornado',
    d:'A cota foi reembolsada. Reúne todos os estornos, independentemente da situação no Totvs. Confirmar se é desistência.' },
  { id:'so_totvs', chip:'FORA DO BOLSÃO', k:'info',
    t:'Matriculado no Totvs, fora do bolsão',
    d:'Matrícula 2027 no Totvs sem pagamento localizado e sem entrada na planilha do bolsão.' },
  { id:'pre_matriculado', chip:'PRÉ-MATRÍCULA', k:'warn',
    t:'Pré-matriculado no Totvs',
    d:'Contrato assinado e vaga reservada, mas a matrícula ainda não foi fechada.' },
  { id:'matriculado', chip:'MATRICULADO', k:'ok',
    t:'Matrícula efetivada',
    d:'Situação Matriculado ou Reserva de vaga - Matriculado no Totvs. É o que conta como matrícula.' }
];

/* Erros de lançamento no financeiro. São marcas, não estados: o aluno continua
   matriculado e o erro aparece ao lado. Só valem para quem já fechou ou
   pré-fechou a matrícula — antes disso não há contrato a cobrar. */
const ERROS = [
  { id:'cota_duplicada', chip:'2 PRIMEIRAS COTAS', k:'crit',
    t:'Aluno com duas primeiras cotas',
    d:'O padrão é 1 cota e 12 mensalidades. Aqui há a reserva de vaga do bolsão e a 1ª cota do contrato, as duas vivas. Fica a que recebeu o pagamento; a outra sai do Totvs. Cancelada, ou zerada com bolsa de 100% e baixada, já conta como resolvida e não aparece.' },
  { id:'sem_plano_ano', chip:'SEM CONTRATO', k:'crit',
    t:'Matrícula fechada sem o plano do ano',
    d:'A entrada foi cobrada e as mensalidades não foram lançadas. O contrato do ano não existe no financeiro: não gera boleto e não entra na receita.' },
  { id:'sem_financeiro', chip:'SEM FINANCEIRO', k:'crit',
    t:'Sem nenhum lançamento financeiro',
    d:'Há matrícula no Totvs e nenhuma cobrança gerada — nem entrada, nem mensalidade.' },
  { id:'baixa_nao_revertida', chip:'REVERTER BAIXA', k:'crit',
    t:'Estornado na Layers e ainda baixado no Totvs',
    d:'A venda foi estornada no e-commerce e a parcela continua baixada no Totvs. O sistema mostra a cota como paga e ela não está — a baixa precisa ser revertida.' },
  { id:'sem_bolsa', chip:'BOLSA NÃO LANÇADA', k:'warn',
    t:'Mensalidade cheia de tabela, sem bolsa',
    d:'A ficha financeira não traz nenhuma dedução nem bolsa cadastrada, então a anuidade entra pelo valor de tabela. Pode ser aluno pagante integral, mas quase sempre é bolsa que a unidade ainda não lançou — e distorce o ticket para cima.' },
  { id:'ticket_fora', chip:'VALOR ATÍPICO', k:'warn',
    t:'Anuidade fora da faixa esperada',
    d:'A anuidade calculada ficou fora do intervalo de R$ 1.800 a R$ 40.000. Fora dessa faixa costuma ser separador decimal errado no lançamento — mas pode ser anuidade real alta. Enquanto não se confirma, o aluno fica fora do cálculo do ticket.' },
  { id:'sem_baixa', chip:'FALTA BAIXAR', k:'warn',
    t:'Pago no e-commerce e ainda não baixado no Totvs',
    d:'A família pagou pela Layers e o dinheiro entrou, mas a parcela continua aberta no Totvs. A conciliação é manual: enquanto não for feita, o sistema mostra dívida que não existe e a cobrança pode ser reenviada à família.' }
];

function errosDe(p, semFicha){
  const t = p.totvs;
  const e = [];
  // A falta de baixa vale para qualquer aluno que pagou, esteja em que
  // situação estiver — o dinheiro entrou de qualquer jeito.
  // Sem a ficha financeira não se sabe o que foi baixado: nenhum erro
  // financeiro pode ser afirmado.
  if(semFicha) return e;
  const pagouLayers = !!(p.mkt && p.mkt.status === 'Pago');
  if(pagouLayers && !(p.fin && p.fin.pago)) e.push('sem_baixa');
  if(p.mkt && p.mkt.status === 'Estornado' && p.fin && p.fin.pago) e.push('baixa_nao_revertida');
  if(!t || !(MATRICULADO.has(t.sit) || PRE.has(t.sit))) return e;
  if(!p.cob || (!p.cob.entradas.length && !p.cob.mensalidades)) e.push('sem_financeiro');
  else {
    if(p.cob.dobra) e.push('cota_duplicada');
    if(p.cob.entradas.length && !p.cob.mensalidades) e.push('sem_plano_ano');
  }
  return e;
}

/* Duas primeiras cotas. O padrão é 1 cota e 12 mensalidades; a dobra é ter a
   reserva de vaga do bolsão E a 1ª cota do contrato, ou a mesma cobrança
   repetida no mesmo vencimento. Parcelas da mesma entrada em vencimentos
   diferentes são uma cota só, partida — não contam.
   Está resolvido, e sai da lista, quando uma das duas recebeu o pagamento e a
   outra foi zerada com bolsa de 100% e baixada sem valor. */
function dobraDe(cob){
  const ent = cob.entradas;
  if(ent.length < 2) return null;
  const tipos = [...new Set(ent.map(x => x.tipo))];
  const repetida = ['reserva','cota'].some(t => {
    const vs = ent.filter(x => x.tipo === t).map(x => x.venc);
    return new Set(vs).size < vs.length;
  });
  if(tipos.length < 2 && !repetida) return null;
  const pagas = ent.filter(x => x.baixado);
  const sobra = ent.filter(x => !x.baixado);
  if(pagas.length && sobra.length && sobra.every(x => x.zerada)) return null;   // resolvido
  // O que sobra é o que a família não deve: as cobranças abertas do tipo que
  // não recebeu dinheiro. A parcela ainda aberta de uma entrada partida que já
  // começou a ser paga não entra — é a mesma cota. Sem dinheiro em nenhuma,
  // fica de pé a reserva.
  const tipoPago = new Set(pagas.map(x => x.tipo));
  const manter = pagas.length ? null : (ent.find(x => x.tipo === 'reserva') || ent[0]);
  const abertas = sobra.filter(x => !x.zerada && x !== manter &&
    (pagas.length ? !tipoPago.has(x.tipo) : x.tipo !== manter.tipo));
  return { valor: abertas.reduce((s, x) => s + x.valor, 0) };
}

function categoria(p, semPlanilha){
  const t = p.totvs;
  // Pagou se entrou pelo Marketplace OU pelo Portal do aluno (Totvs).
  const pago = (p.mkt && p.mkt.status==='Pago') || (p.fin && p.fin.pago);
  // pagamento desfeito tem precedência: não pode contar como matrícula fechada
  // Todo estorno vai para a mesma lista, seja qual for a situação no Totvs:
  // dinheiro devolvido exige decisão antes de qualquer outra coisa.
  if(p.mkt && p.mkt.status==='Estornado') return 'estorno';
  if(p.mkt && p.mkt.status==='Pendente / Vencido' && !(p.fin && p.fin.pago)){
    if(!t || !avancou(t.sit)) return 'pgto_pendente';
  }
  if(t){
    // "Fora do bolsão" só faz sentido quando existe a planilha para comparar.
    // Sem ela, a situação do Totvs manda sozinha.
    const noBolsao = pago || p.plan || semPlanilha;
    if(MATRICULADO.has(t.sit)) return noBolsao ? 'matriculado' : 'so_totvs';
    if(PRE.has(t.sit))         return noBolsao ? 'pre_matriculado' : 'so_totvs';
    if(PENDENTE.has(t.sit))    return 'totvs_pendente';
    return 'totvs_cancel';
  }
  if(pago) return 'falta_totvs';
  // O registro sem FLUXO preenchido é o mesmo problema: o aluno está só na
  // planilha e nada aconteceu. A coluna em branco vira uma marca na lista.
  return 'falta_pagamento';
}

/* ---------------- montagem ---------------- */

function cruzar(caminhoTotvs, caminhoUnidades, caminhoMkt, caminhoFin){
  const tv = lerTotvs(caminhoTotvs);
  const un = lerUnidades(caminhoUnidades);
  const mk = lerMarketplace(caminhoMkt);
  const fin = lerFinanceiro(caminhoFin);

  // Um aluno pode ter mais de uma transação (estornou e pagou de novo). Junto
  // antes de parear: senão a linha estornada consome o registro do Totvs e a
  // paga fica órfã, aparecendo como "não está no Totvs".
  juntarRepetidos(mk.pessoas);

  // Marketplace ↔ planilha
  parear(mk.pessoas, un, 'aluno', 'aluno');
  // (Marketplace ou planilha) ↔ Totvs
  const pessoas = [];
  mk.pessoas.forEach(g => pessoas.push({ code:g.code, mkt:g, plan:g._par||null, totvs:null,
    nomes:[g.aluno, g._par&&g._par.aluno].filter(Boolean) }));
  un.forEach(u => { if(!u._par) pessoas.push({ code:u.code, mkt:null, plan:u, totvs:null, nomes:[u.aluno] }); });

  // Levo o responsável junto: é ele que confirma o par quando o nome do aluno
  // vem escrito diferente no Marketplace e no Totvs.
  const alvo = pessoas.map(p => ({ code:p.code, aluno:p.nomes[0], _p:p,
    resp: (p.mkt && p.mkt.resp) || (p.plan && p.plan.resp) || null }));
  parear(alvo, tv.captacao, 'aluno', 'aluno');
  alvo.forEach(a => { if(a._par) a._p.totvs = a._par; a._p._notaTv = a._nota||null; });
  // quem sobrou no Totvs entra como aluno de fora do bolsão
  tv.captacao.forEach(t => { if(!t._par) pessoas.push({ code:t.code, mkt:null, plan:null, totvs:t, nomes:[t.aluno] }); });

  // O financeiro casa por RA — chave exata, sem depender de nome.
  pessoas.forEach(p => { p.fin = (p.totvs && p.totvs.ra) ? (fin.porRA.get(p.totvs.ra) || null) : null; });
  pessoas.forEach(p => {
    p.cob = (p.totvs && p.totvs.ra) ? (fin.cobranca.get(p.totvs.ra) || null) : null;
    if(!p.cob) return;
    const d = dobraDe(p.cob);
    p.cob.dobra = !!d;
    p.cob.duplicado = d ? d.valor : 0;
  });
  // RA anterior a 2027 sem rastro de bolsão — nem pedido na Layers, nem reserva
  // de vaga no Totvs — é aluno da casa lançado como matrícula nova. Sai da
  // captação e vai para o aviso de qualidade, para a unidade corrigir o tipo.
  const semRastro = p => p.totvs && p.totvs.ra && !/^MT27/i.test(String(p.totvs.ra)) && !p.mkt &&
    !(p.cob && p.cob.entradas.some(e => e.tipo === 'reserva'));
  tv.relato.tipoAConferir = pessoas.filter(semRastro).map(p => ({ aluno: p.totvs.aluno, ra: p.totvs.ra,
    sit: p.totvs.sit, serie: p.totvs.serie, unidade: CODE[p.code] || p.code, motivo: 'RA de ' + String(p.totvs.ra).slice(2,4) + ' sem bolsão' }));
  const captacao = pessoas.filter(p => !semRastro(p));
  pessoas.length = 0; captacao.forEach(p => pessoas.push(p));

  const semPlanilha = un.length === 0;
  pessoas.forEach(p => { p.cat = categoria(p, semPlanilha); p.erros = errosDe(p, !!fin.relato.ausente); });

  const dias = iso => {
    if(!iso) return null;
    const hoje = new Date(); const h = Date.UTC(hoje.getFullYear(),hoje.getMonth(),hoje.getDate());
    const q = iso.split('-');
    return Math.max(0, Math.round((h - Date.UTC(+q[0],+q[1]-1,+q[2]))/86400000));
  };

  const itens = pessoas.map(p => {
    const g=p.mkt, u=p.plan, t=p.totvs;
    const mf = (t && t.ra) ? (fin.mensPorRA && fin.mensPorRA.get(t.ra)) || null : null;
    const dISO = g ? isoDe(g.data) : null;
    return {
      code:p.code, unidade:CODE[p.code]||p.code, cat:p.cat,
      nome: (g&&g.aluno) || (t&&t.aluno) || (u&&u.aluno),
      serie: (t&&t.serie) || (g&&g.seriePret) || (u&&u.turma) || '',
      data: g?g.data:'', dISO, dias: dias(dISO),
      bolsao: dISO ? (dISO>='2026-08-22'?2:1) : 0,
      pg: g?g.status:'', valor: g?g.valor:null, pedido: g?g.pedido:null,
      // origem do pagamento: M = Marketplace, P = Portal do aluno (Totvs)
      pagoM: !!(g && g.status==='Pago'),
      // Venda estornada não é cota paga, ainda que a parcela siga baixada no
      // Totvs: nesse caso o que falta é reverter a baixa, não cobrar.
      pagoP: !!(p.fin && p.fin.pago) && !(g && g.status==='Estornado'),
      baixaNaoRevertida: !!(g && g.status==='Estornado' && p.fin && p.fin.pago),
      parcial: !!(p.fin && p.fin.parcial),
      falta: p.fin && p.fin.parcial ? p.fin.parcial.falta : null,
      cotaEsperada: p.fin ? p.fin.esperado : null,
      finDt: p.fin?p.fin.dtBaixa:null, finValor: p.fin?p.fin.valor:null,
      finVenc: p.fin && !p.fin.pago ? p.fin.venc : null,
      trilha: p.fin?p.fin.descricao:null,
      // Erros de lançamento no financeiro, para as listas de rastreio.
      nEntradas: p.cob ? p.cob.entradas.length : 0,
      entradas: p.cob ? p.cob.entradas : [],
      temMensalidade: p.cob ? p.cob.mensalidades > 0 : false,
      duplicado: p.cob ? p.cob.duplicado : 0,
      erros: p.erros || [],
      resp: g?g.resp:(t?t.respFin:null), tel:g?g.tel:null, email:g?g.email:null,
      plan: !!u, fluxo: u?u.fluxo:null, linha: u?u.linha:null,
      ticket: u?u.ticket:null, mensalidade: u?u.mensalidade:null, bolsa: u?u.bolsa:null,
      // Anuidade pela ficha financeira: soma das mensalidades mais a cota da
      // unidade. Mesma régua da renovação, e não depende da planilha.
      ticketFin: mf ? mf.anoMens + (p.code === 'AM' ? COTA_RECREIO : COTA_PADRAO) : null,
      mensalFin: mf ? mf.mensal : null, semBolsaFin: mf ? !!mf.semBolsa : false,
      ra: t?t.ra:null, sit: t?t.sit:'', turma: t?t.turma:null, grade: t?t.grade:null,
      matriculado: !!(t && MATRICULADO.has(t.sit)),
      pre: !!(t && PRE.has(t.sit))
    };
  });

  // liga o ticket da renovação pelo RA
  tv.renovacao.forEach(a => {
    const m = a.ra ? fin.mensPorRA && fin.mensPorRA.get(a.ra) : null;
    a.mensal = m ? m.mensal : null;
    a.nParc  = m ? m.nParc  : null;
    a.somaParc = m ? m.soma : null;
    // A cota segue a unidade onde o aluno está matriculado, não a filial que
    // lançou a parcela: há aluno de uma unidade com lançamento em outra.
    a.cota   = m ? (a.code === 'AM' ? COTA_RECREIO : COTA_PADRAO) : null;
    a.anual  = m ? m.anoMens + a.cota : null;
    a.filialFin = m ? m.filial : null;
    a.semBolsa = m ? m.semBolsa : false;
  });

  return { CODE, ORDER, itens, totvs: tv, unidades: un, mkt: mk, fin };
}

module.exports = { cruzar, CATEGORIAS, ERROS, CODE, ORDER, MATRICULADO, PRE, PENDENTE, ENCERRADO, parecenca, SUGERE };

if(require.main === module){
  const [a,b,c,e] = process.argv.slice(2);
  if(!a||!b){ console.error('uso: node cruzar.js <totvs.xlsx> <unidades.xlsx> [marketplace.xlsx]'); process.exit(1); }
  const r = cruzar(a,b,c,e);
  const cont = {}; r.itens.forEach(i => cont[i.cat]=(cont[i.cat]||0)+1);
  console.error('Totvs:', JSON.stringify(r.totvs.relato));
  console.error('Marketplace:', JSON.stringify(r.mkt.relato));
  console.error('financeiro:', JSON.stringify(r.fin.relato));
  console.error('planilha das unidades:', r.unidades.length, 'registros');
  console.error('captação:', r.itens.length, 'alunos');
  CATEGORIAS.forEach(c2 => console.error('  '+String(cont[c2.id]||0).padStart(4), c2.t));
  console.error('renovação: elegíveis', r.totvs.renovacao.length,
    '| matriculados', r.totvs.renovacao.filter(x=>x.matriculado).length,
    '| pré', r.totvs.renovacao.filter(x=>x.pre).length);
  process.stdout.write(JSON.stringify(r.itens));
}
