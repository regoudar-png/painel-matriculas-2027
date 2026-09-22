// Monta o JSON que o painel consome, a partir do cruzamento das três bases.
//
//   node motor/montar-painel.js <totvs.xlsx> <unidades.xlsx> [marketplace.xlsx] [--mkt-ate AAAA-MM-DD]
//
// Escreve o snapshot em dados/painel.json.

const fs = require('fs');
const path = require('path');
const { cruzar, CATEGORIAS, ERROS, CODE, ORDER, MATRICULADO, PRE, parecenca, SUGERE } = require('./cruzar.js');

const PREVISTO = 3100;   // meta de captação do ciclo 2027

// Faixa plausível de anuidade. Fora dela quase sempre é separador decimal
// digitado errado na planilha — deixar entrar destrói a média.
const TICKET_MIN = 1800, TICKET_MAX = 40000;
const ticketOk = v => v != null && v >= TICKET_MIN && v <= TICKET_MAX;
// A ficha financeira manda; a planilha das unidades entra só onde a ficha não
// alcança, e pode nem existir.
const ticketDe = i => i.ticketFin != null ? i.ticketFin : i.ticket;
// Mediana, não média: aluno com a bolsa não lançada entra pelo valor cheio de
// tabela e desloca a média inteira de uma unidade pequena.
const mediana = v => { if(!v.length) return null;
  const s = [...v].sort((a,b)=>a-b), n = s.length;
  return n % 2 ? s[(n-1)/2] : (s[n/2-1] + s[n/2]) / 2; };

/* Todo o cálculo do painel a partir do cruzamento já feito. Recebe o retorno
   de cruzar() e a data do último export do Marketplace; devolve o JSON que a
   página consome. Não toca em disco — o navegador chama a mesma função. */
function montar(r, mktAte, fontes){
  fontes = fontes || {};
  const { itens, totvs } = r;

  /* ---- marca quem entrou depois do último retrato do Marketplace ---- */
  const dtPorAluno = new Map(totvs.captacao.map(t => [t.aluno, t.dtMat]));
  let defasados = 0;
  itens.forEach(i => {
    const dt = dtPorAluno.get(i.nome);
    i.dtMat = dt || null;
    i.posMkt = !!(mktAte && dt && dt > mktAte && !i.pg);
    if(i.posMkt) defasados++;
  });

  /* ---- listas por categoria ---- */
  const listas = CATEGORIAS.map(c => {
    const linhas = itens.filter(i => i.cat === c.id);
    const porUni = {};
    linhas.forEach(i => porUni[i.code] = (porUni[i.code]||0)+1);
    return { ...c, n: linhas.length, porUni,
      itens: linhas.map(i => ({
        u:i.code, nome:i.nome, serie:i.serie, data:i.data, dias:i.dias, dISO:i.dISO,
        bolsao:i.bolsao, pg:i.pg, pagoM:i.pagoM, pagoP:i.pagoP, parcial:i.parcial, falta:i.falta,
        cotaEsperada:i.cotaEsperada,
        finDt:i.finDt, finVenc:i.finVenc, baixaZero:!!i.baixaZero,
        trilha:i.trilha, plan:i.plan, ra:i.ra||null, sit:i.sit,
        matriculado:i.matriculado, pre:i.pre, ticket:i.ticket, posMkt:i.posMkt, sug:i.sug||null,
        semFluxo: i.fluxo === '(em branco)'
      })).sort((a,b) => (a.dISO||'9').localeCompare(b.dISO||'9') ||
        ORDER.indexOf(a.u)-ORDER.indexOf(b.u) || a.nome.localeCompare(b.nome,'pt'))
    };
  });
  const cnt = id => (listas.find(l => l.id===id)||{n:0}).n;

  /* ---- contadores do topo ---- */
  // Cada caixa é um recorte que não se sobrepõe aos outros na conta final:
  // matriculados + pré + em exigência + descartados = total da captação.
  const CONCLUIDO = ['matriculado','pre_matriculado'];
  const DESCARTADO = ['totvs_cancel','so_totvs'];
  const pagoM = itens.filter(i => i.pagoM).length;
  // Na caixa aparece o que a Layers mostra: linhas com status Pago.
  const linhasPagasMkt = r.mkt.relato.linhasPagas || pagoM;
  const pagoP = itens.filter(i => i.pagoP && !i.pagoM).length;
  const naPlanilha = itens.filter(i => i.plan).length;
  const emExigencia = itens.filter(i => !CONCLUIDO.includes(i.cat) && !DESCARTADO.includes(i.cat));
  const motivos = {};
  emExigencia.forEach(i => motivos[i.cat] = (motivos[i.cat]||0)+1);
  const maiorMotivo = Object.entries(motivos).sort((a,b)=>b[1]-a[1])[0];
  const nomeCat = id => (CATEGORIAS.find(c=>c.id===id)||{t:id}).t;

  // As caixas seguem a jornada da matrícula, na ordem em que ela acontece. Entre
  // duas caixas fica o vão: quem chegou na etapa anterior e não passou para a
  // seguinte. Assim a diferença entre os números fica explícita, e cada vão é uma
  // lista clicável — não é subtração de cabeça.
  // Sem a planilha das unidades não há o que contar nessa etapa: a caixa e o
  // vão dela saem da faixa, em vez de aparecerem zerados — zero ali leria como
  // "ninguém registrado", que é falso.
  const temPlanilha = r.unidades.length > 0;
  const kpi = [
    temPlanilha ? {k:'info', sel:{tipo:'plan'}, lb:'Na planilha da unidade', v: naPlanilha,
     sb:'registrados pelas secretarias'} : null,
    {k:'ok', sel:{tipo:'pago'}, lb:'1ª cota paga', v: pagoM + pagoP,
     sb: pagoM+' Marketplace · '+pagoP+' Portal',
     partes:[{cor:'var(--ok)', v:pagoM, rot:'Marketplace'},
             {cor:'var(--info)', v:pagoP, rot:'Portal do aluno'}]},
    {k:'info', sel:{tipo:'totvs'}, lb:'Lançados no Totvs', v: itens.filter(i=>i.ra).length,
     sb:'já têm RA para 2027'},
    {k:'ok', sel:{tipo:'cat', ids:['matriculado']}, lb:'Matriculados', v: itens.filter(i=>i.cat==='matriculado').length,
     sb:'situação Matriculado no Totvs'},
    {k:'warn', sel:{tipo:'cat', ids:['pre_matriculado']}, lb:'Pré-matriculados', v: itens.filter(i=>i.cat==='pre_matriculado').length,
     sb:'contrato assinado, falta fechar'}
  ].filter(Boolean);

  // Cada vão fica entre a caixa de índice `apos` e a seguinte.
  const VAO = {
    plan_sem_pago:    i => i.plan && !(i.pagoM || i.pagoP),
    pago_sem_totvs:   i => (i.pagoM || i.pagoP) && !i.ra,
    // Pela situação no Totvs, não pela categoria: há aluno Matriculado no Totvs
    // que caiu em outra lista por não ter pagamento localizado. Ele fechou.
    totvs_sem_fechar: i => i.ra && !i.matriculado && !i.pre
  };
  // O índice de cada vão acompanha a caixa anterior, que anda uma casa para
  // trás quando a caixa da planilha sai de cena.
  const d = temPlanilha ? 0 : 1;
  const vaos = [
    temPlanilha ? { apos:0, f:'plan_sem_pago', curto:'sem pagar',
      lb:'estão na planilha da unidade e ainda não pagaram a 1ª cota' } : null,
    { apos:1-d, f:'pago_sem_totvs',   curto:'falta lançar',
      lb:'pagaram a 1ª cota e a unidade ainda não lançou a matrícula no Totvs' },
    { apos:2-d, f:'totvs_sem_fechar', curto:'falta concluir',
      lb:'já têm RA, mas a situação no Totvs ainda é Pendente, Cancelado ou Desistente' }
  ].filter(Boolean).map(v => ({ ...v, v: itens.filter(VAO[v.f]).length }));

  /* ---- placar por unidade, ordenado por matriculados ---- */
  const media = xs => xs.length ? xs.reduce((a,b)=>a+b,0)/xs.length : null;
  const placar = ORDER.map(c => {
    const it = itens.filter(i => i.code===c);
    const g = id => it.filter(i => i.cat===id).length;
    const tickets     = it.map(ticketDe).filter(ticketOk);
    const ticketsMat  = it.filter(i=>i.cat==='matriculado').map(ticketDe).filter(ticketOk);
    return { code:c, nome:CODE[c], tot:it.length,
      mat:g('matriculado'), pre:g('pre_matriculado'),
      falta:g('falta_totvs'),
      pend:g('totvs_pendente'), cancel:g('totvs_cancel'),
      semPg:g('falta_pagamento')+g('pgto_pendente')+g('estorno'),
      fora:g('so_totvs'),
      ticket: mediana(tickets), nTicket: tickets.length,
      ticketMat: mediana(ticketsMat), nTicketMat: ticketsMat.length,
      // Quantos dos matriculados da unidade têm anuidade calculável. Sem isso
      // o ticket parece falar por todos, e às vezes fala por metade.
      nMat: it.filter(i=>i.cat==='matriculado').length };
  }).sort((a,b) => b.mat-a.mat || b.pre-a.pre || a.nome.localeCompare(b.nome,'pt'));
  placar.forEach((p,i) => p.pos = i+1);

  /* ---- bolsões ---- */
  const bolsoes = [
    {n:1, rot:'Bolsão 1', sub:'prova 08/08', v: itens.filter(i=>i.bolsao===1).length},
    {n:2, rot:'Bolsão 2', sub:'prova 22/08', v: itens.filter(i=>i.bolsao===2).length},
    {n:0, rot:'Portal ou sem pagamento', sub:'fora dos bolsões', v: itens.filter(i=>!i.bolsao).length}
  ];

  // A renovação saiu do painel: o foco é captação.

  /* ---- possíveis correspondências ainda não unidas ---- */
  const semTotvs = itens.filter(i => !i.ra && (i.pg==='Pago' || i.plan));
  const sobrouTotvs = totvs.captacao.filter(t => !t._par);
  const sugestoes = [];
  semTotvs.forEach(i => {
    sobrouTotvs.filter(t => t.code===i.code).forEach(t => {
      const s = parecenca(i.nome, t.aluno);
      if(s>=SUGERE && s<0.80) sugestoes.push({ u:i.code, a:i.nome, b:t.aluno, s:Math.round(s*100), sit:t.sit });
    });
  });
  sugestoes.sort((a,b)=>b.s-a.s);
  // A melhor sugestão de cada aluno vai junto com ele, para aparecer na lista.
  const melhorSug = new Map();
  sugestoes.forEach(g => {
    const k = g.u + '|' + g.a;
    if(!melhorSug.has(k) || melhorSug.get(k).s < g.s) melhorSug.set(k, g);
  });
  itens.forEach(i => {
    const g = melhorSug.get(i.code + '|' + i.nome);
    i.sug = g ? { nome: g.b, s: g.s, sit: g.sit } : null;
  });

  const semPagamento = itens.filter(i => !i.pagoM && !i.pagoP)
    .map(i => ({ u:i.code, nome:i.nome, sit:i.sit||"fora do Totvs", plan:i.plan, cat:i.cat }));

  const ticketFora = itens.filter(i => ticketDe(i) != null && !ticketOk(ticketDe(i)))
    .map(i => ({ u: i.code, nome: i.nome, ticket: ticketDe(i) }));

  // Bolsa não lançada e valor atípico só se sabe aqui, onde a faixa de
  // plausibilidade é aplicada.
  itens.forEach(i => {
    if(!i.matriculado && !i.pre) return;
    const v = ticketDe(i);
    if(i.semBolsaFin && ticketOk(v)) i.erros = (i.erros || []).concat('sem_bolsa');
    if(v != null && !ticketOk(v)) i.erros = (i.erros || []).concat('ticket_fora');
  });

  /* ---- erros de lançamento: marcas, não estados ---- */
  // O mesmo aluno pode estar em mais de uma lista, e continua contando como
  // matriculado. É isso que permite acompanhar se a unidade zerou o problema.
  const erros = ERROS.map(e => {
    const linhas = itens.filter(i => (i.erros || []).includes(e.id));
    const porUni = {};
    linhas.forEach(i => porUni[i.code] = (porUni[i.code] || 0) + 1);
    return { ...e, n: linhas.length, porUni,
      ras: linhas.map(i => i.ra).filter(Boolean),
      itens: linhas.map(i => ({
        u: i.code, nome: i.nome, serie: i.serie, ra: i.ra || null, sit: i.sit,
        nEntradas: i.nEntradas || 0, duplicado: i.duplicado || 0,
        temMensalidade: !!i.temMensalidade, anual: ticketDe(i),
        entradas: (i.entradas || []).map(x => ({ servico: x.servico, venc: x.venc,
          valor: x.valor, baixado: x.baixado, zerada: !!x.zerada }))
      })).sort((a, b) => ORDER.indexOf(a.u) - ORDER.indexOf(b.u) ||
        a.nome.localeCompare(b.nome, 'pt'))
    };
  });

  const saida = {
    CODE, ORDER, previsto: PREVISTO,
    geradoEm: new Date().toISOString().slice(0,10),
    fontes: {
      totvs: fontes.totvs || null,
      unidades: fontes.unidades || null,
      marketplace: fontes.marketplace || null,
      financeiro: fontes.financeiro || null,
      mktAte, defasados
    },
    erros,
    cap: { kpi, vaos, listas, placar, bolsoes,
      exigencia: { total: emExigencia.length, motivos, ids: Object.keys(motivos) },
      previsto: PREVISTO,
      total: itens.length,
      matriculados: cnt('matriculado'),
      pre: cnt('pre_matriculado'),
      pagos: itens.filter(i=>i.pagoM||i.pagoP).length,
      pagosM: itens.filter(i=>i.pagoM).length,
      pagosP: itens.filter(i=>i.pagoP && !i.pagoM).length,
      ticket: mediana(itens.map(ticketDe).filter(ticketOk)),
      ticketMat: mediana(itens.filter(i=>i.cat==='matriculado').map(ticketDe).filter(ticketOk)),
      semBolsa: itens.filter(i => i.matriculado && i.semBolsaFin)
        .map(i => ({ u:i.code, nome:i.nome, serie:i.serie, anual:i.ticketFin }))
        .sort((a,b) => (b.anual||0) - (a.anual||0)) },
    qualidade: { ...totvs.relato, financeiro: r.fin.relato, marketplace: r.mkt.relato, sugestoes: sugestoes.slice(0,20),
      ticketFora, semPagamento }
  };


  // Números que só interessam a quem está importando, não ao painel.
  saida.diag = { defasados, mktAte: mktAte || null,
    nTicketFora: ticketFora.length, nSemPagamento: semPagamento.length,
    nSemBolsa: saida.cap.semBolsa.length,
    topo: placar.slice(0,4).map(p => ({ pos:p.pos, code:p.code, mat:p.mat })) };
  return saida;
}

if (typeof module !== 'undefined' && module.exports) module.exports = { montar, PREVISTO };

/* ---------------- linha de comando ---------------- */
if (typeof require !== 'undefined' && require.main === module) {
  const argv = process.argv.slice(2);
  const flags = {}, pos = [];
  for(let i=0;i<argv.length;i++){
    if(argv[i].startsWith('--')) flags[argv[i].slice(2)] = argv[++i];
    else pos.push(argv[i]);
  }
  // Os arquivos são reconhecidos pelo conteúdo, não pela ordem — o export da
  // Layers mudou de formato e a planilha das unidades virou opcional, então
  // posição de argumento vira armadilha.
  const { lerPlanilha: _ler } = require('./ler-xlsx.js');
  const assinatura = p => {
    let wb; try { wb = _ler(p); } catch(e) { return null; }
    const abas = wb.abas.filter(a => a.linhas && a.linhas.length);
    if(!abas.length) return null;
    const maior = abas.slice().sort((a,b) => b.linhas.length - a.linhas.length)[0];
    const cab = (maior.linhas[0]||[]).map(v => String(v==null?'':v).trim().toUpperCase());
    const tem = (...ks) => ks.every(k => cab.indexOf(k) >= 0);
    if(tem('VALOR_BAIXADO','SERVICO','RA')) return 'fin';
    if(tem('SITUACAO MATRICULA','TIPO MATRICULA','NOME ALUNO')) return 'totvs';
    if(tem('CÓDIGO DA VENDA','STATUS DA VENDA')) return 'mkt';
    if(tem('STATUS','ID DA TRANSAÇÃO') || tem('STATUS','CÓDIGO DO PEDIDO')) return 'mkt';
    if(tem('IDLAN','PLANO DE PAGAMENTO','BOLSA %')) return 'bolsas';   // reconhecido, não usado
    return null;
  };
  const achado = {};
  pos.forEach(p => { const k = assinatura(p);
    if(k === 'bolsas') console.error('ignorado (o painel não usa o relatório de bolsas):', path.basename(p));
    else if(k && !achado[k]) achado[k] = p;
    else if(!k) console.error('não reconheci:', path.basename(p)); });
  const pTotvs = achado.totvs, pUnidades = achado.unidades || null,
        pMkt = achado.mkt || null, pFin = achado.fin || null;
  console.error('fontes:', Object.entries(achado).map(([k,v]) => k+'='+path.basename(v)).join(' | '));
  if(!pTotvs){
    console.error('uso: node montar-painel.js <arquivos.xlsx...> [--mkt-ate AAAA-MM-DD]');
    console.error('a base de alunos do Totvs é obrigatória; os outros são opcionais');
    process.exit(1);
  }
  // Data do export do Marketplace. Tudo que entrou no Totvs depois disso pode
  // ter pagamento que a base ainda não enxerga.
  const ate = flags['mkt-ate'] || ((pMkt||'').match(/(\d{4}-\d{2}-\d{2})/)||[])[1] || null;

  const nome = p => p ? path.basename(p) : null;
  const saida = montar(cruzar(pTotvs, pUnidades, pMkt, pFin), ate,
    { totvs:nome(pTotvs), unidades:nome(pUnidades), marketplace:nome(pMkt), financeiro:nome(pFin) });
  const d = saida.diag;
  const f2 = v => v == null ? '—' : v.toFixed(2);

  const destino = path.join(__dirname, '..', 'dados', 'painel.json');
  fs.mkdirSync(path.dirname(destino), { recursive: true });
  fs.writeFileSync(destino, JSON.stringify(saida));

  console.log('captação:', saida.cap.total, 'alunos ·', saida.cap.matriculados, 'matriculados ·',
    saida.cap.pre, 'pré');
  console.log('ticket médio geral:', f2(saida.cap.ticket), '| dos matriculados:', f2(saida.cap.ticketMat));
  saida.erros.forEach(e => console.log('  ' + String(e.n).padStart(4), e.t));
  if(d.defasados) console.log('ATENÇÃO:', d.defasados, 'alunos entraram no Totvs depois de', d.mktAte,
    '— o pagamento pode existir e não estar na base do Marketplace');
  if(d.nTicketFora) console.log(String(d.nTicketFora), 'tickets fora da faixa:',
    saida.qualidade.ticketFora.map(t => t.u + ' ' + t.nome.slice(0,22)).join(' | '));
  console.log('pagamento: Marketplace', saida.cap.pagosM, '| Portal do aluno', saida.cap.pagosP,
    '| sem pagamento localizado', d.nSemPagamento);
  console.log('ranking:', d.topo.map(p=>p.pos+'º '+p.code+' ('+p.mat+')').join('  '));
  console.log('->', destino, Math.round(fs.statSync(destino).size/1024)+' KB');
}
