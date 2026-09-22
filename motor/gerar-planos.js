const fs=require('fs'),path=require('path');
const {pasta,colName}=require('./xlsxw.js');
const D=require('./final.json');
const ITEMS=D.items, CODE=D.CODE;
const ORDER=['BG','CG','AM','CX','RM','TQ','TJ','NI','MD','RT','SJ'];
const SAIDA=process.argv[2]||'.';

/* ---- regras dos planos ---- */
const pago=i=>i.statusGw==='Pago';
const semMkt=i=>!i.statusGw;
const foraTotvs=i=>!i.ra;
const PLANOS=[
 {n:1,aba:'1. Baixar 1ª cota',
  criterio:'Marketplace “pago” + Totvs “Pré-Matriculado”',
  cond:i=>pago(i)&&i.tvSit==='Pré-Matriculado',
  acao:'Dar baixa na 1ª cota no Totvs para converter em matrícula.'},
 {n:2,aba:'2. Cadastrar no Totvs',
  criterio:'Marketplace “pago” + fora do Totvs',
  cond:i=>pago(i)&&foraTotvs(i),
  acao:'Cadastrar o aluno no Totvs e, em seguida, dar baixa na 1ª cota para converter em matrícula.'},
 {n:3,aba:'3. Cadastrar e cobrar',
  criterio:'Sem pedido no Marketplace + consta na planilha da unidade + fora do Totvs',
  cond:i=>semMkt(i)&&i.alunoUn&&foraTotvs(i),
  acao:'Cadastrar o aluno no Totvs e cobrar o pagamento da 1ª cota.',
  justifica:true},
 {n:4,aba:'4. Assinar contrato',
  criterio:'Totvs “Reserva de vaga - Matriculado”',
  cond:i=>i.tvSit==='Reserva de vaga - Matriculado',
  acao:'Contatar a família para assinatura do contrato de prestação de serviço e, depois, virar para matriculado no Totvs.'},
 {n:5,aba:'5. Contrato (reserva pendente)',
  criterio:'Marketplace “pago” + Totvs “Reserva de vaga - Pendente”',
  cond:i=>pago(i)&&i.tvSit==='Reserva de vaga - Pendente',
  acao:'Contatar a família para assinatura do contrato de prestação de serviço e, depois, virar para matriculado no Totvs.'},
 {n:6,aba:'6. Contrato e baixa',
  criterio:'Marketplace “pago” + Totvs “Pendente”',
  cond:i=>pago(i)&&i.tvSit==='Pendente',
  acao:'Contatar a família para assinatura do contrato e, em seguida, encaminhar a baixa da 1ª cota no Totvs para converter em matrícula.'}
];
// Regras 7 a 12 fecham os casos que as seis primeiras não alcançavam.
// A lógica que as une: 'Pré-Matriculado' e 'Reserva de vaga - Pré-Mat' significam
// contrato assinado — se pagou, basta baixar a cota; se não pagou, é cobrança.
// 'Pendente' e 'Reserva de vaga - Pendente' significam contrato ainda não assinado.
const PRE=s=>s==='Pré-Matriculado'||s==='Reserva de vaga - Pré-Mat';
const PEND=s=>s==='Pendente'||s==='Reserva de vaga - Pendente';
PLANOS.push(
 {n:7,aba:'7. Baixar 1ª cota (reserva)',
  criterio:'Marketplace “pago” + Totvs “Reserva de vaga - Pré-Mat”',
  cond:i=>pago(i)&&i.tvSit==='Reserva de vaga - Pré-Mat',
  acao:'Dar baixa na 1ª cota no Totvs para converter em matrícula.'},
 {n:8,aba:'8. Contrato e pagamento',
  criterio:'Sem pagamento confirmado + Totvs “Pendente” ou “Reserva de vaga - Pendente”',
  cond:i=>!pago(i)&&PEND(i.tvSit),
  acao:'Contatar a família para finalizar a matrícula online: assinar o contrato de prestação de serviço e pagar a 1ª cota.'},
 {n:9,aba:'9. Cobrar 1ª cota',
  criterio:'Sem pagamento confirmado + Totvs “Pré-Matriculado” ou “Reserva de vaga - Pré-Mat”',
  cond:i=>!pago(i)&&PRE(i.tvSit),
  acao:'Contrato já assinado — cobrar o pagamento da 1ª cota e, depois, dar baixa para converter em matrícula.'},
 {n:10,aba:'10. Retomar cobrança',
  criterio:'Fora do Totvs + pagamento vencido ou estornado no Marketplace',
  cond:i=>foraTotvs(i)&&(i.statusGw==='Pendente / Vencido'||i.statusGw==='Estornado'),
  acao:'Confirmar o interesse da família: se mantém, refazer a cobrança da 1ª cota e cadastrar no Totvs; se desistiu, registrar a desistência.',
  extra:['Retorno da família',34]},
 {n:11,aba:'11. Confirmar desistência',
  criterio:'Totvs “Desistente” ou “Cancelado”',
  cond:i=>i.tvSit==='Desistente'||i.tvSit==='Cancelado',
  acao:'Confirmar a desistência com a família, tratar o estorno quando houver e retirar o aluno da contagem de matrículas.',
  extra:['Motivo informado pela família',34]},
 {n:12,aba:'12. Concluído - sem ação',
  criterio:'Totvs “Matriculado”',
  cond:i=>i.tvSit==='Matriculado',
  acao:'Matrícula concluída. Nenhuma ação necessária — a aba serve de conferência.'}
);
const SEMPLANO={n:0,aba:'Sem plano definido',
  criterio:'Não se encaixa em nenhuma das regras',
  acao:'Sem regra definida — conferir com a coordenação antes de agir.'};

ITEMS.forEach(i=>{const p=PLANOS.find(p=>p.cond(i)); i.plano=p?p.n:0;});

/* ---- colunas ---- */
const COLS=[
 ['Aluno',30],['Série',13],['Turma',12],['Responsável',26],['Telefone',14],['E-mail',26],
 ['CPF do responsável',15],['RA no Totvs',13],['Situação no Totvs',20],
 ['Status no Marketplace',16],['Data do pagamento',13],['Dias parado',8],['Bolsão',8],
 ['Valor pago',11],['Código do pedido',17],['Na planilha da unidade',18],
 ['Mensalidade (12x)',13],['Bolsa negociada',9],['Ação necessária',48],
 ['Resolvido?',11],['Quem resolveu',18],['Data da conclusão',13],['Observação',30]];
const COL_JUST=['Justificativa da unidade',44];
const extraDe=p=>p.justifica?COL_JUST:(p.extra||null);
const iResolv=19; // coluna T

const HOJE=(()=>{const n=new Date();return Date.UTC(n.getFullYear(),n.getMonth(),n.getDate());})();
const dias=iso=>{if(!iso)return null;const p=iso.split('-');
  return Math.max(0,Math.round((HOJE-Date.UTC(+p[0],+p[1]-1,+p[2]))/86400000));};
const S={hdr:1,txt:2,data:3,moeda:4,pct:5,num:6,acao:7,marca:8,titulo:9,legenda:10,verde:11};

function linhaAluno(i,plano){
  const c=v=>({v,t:'s',s:S.txt});
  const n=v=>({v:v==null?'':v,t:v==null?'s':'n',s:S.num});
  const cells=[
   c(i.nome),
   c(i.tvSerie||i.seriePret||''),
   c(i.tvTurma||i.turma||''),
   c(i.resp||i.tvResp||''),
   c(i.tel||''),
   c(i.email||''),
   c(i.cpf||i.tvCpfResp||''),
   c(i.ra||''),
   c(i.tvSit||'Fora do Totvs'),
   c(i.statusGw||'Sem pedido'),
   i.dISO?{v:i.dISO,t:'d',s:S.data}:{v:'',t:'s',s:S.data},
   n(dias(i.dISO)),
   c(i.bolsao?'Bolsão '+i.bolsao:''),
   i.valor!=null?{v:i.valor,t:'m',s:S.moeda}:{v:'',t:'s',s:S.moeda},
   c(i.pedido||''),
   c(i.alunoUn?(i.fluxo||'registrado')+' · linha '+i.linha:'Não consta'),
   i.mensalidade!=null?{v:i.mensalidade,t:'m',s:S.moeda}:{v:'',t:'s',s:S.moeda},
   i.bolsa!=null?{v:i.bolsa,t:'p',s:S.pct}:{v:'',t:'s',s:S.pct},
   {v:plano.acao,t:'s',s:S.acao},
   {v:'NÃO',t:'s',s:S.marca},
   {v:'',t:'s',s:S.txt},
   {v:'',t:'s',s:S.data},
   {v:'',t:'s',s:S.txt}];
  if(extraDe(plano))cells.push({v:'',t:'s',s:S.txt});
  return {cells,h:30};
}

function abaPlano(plano,itens){
  const ex=extraDe(plano);
  const cols=ex?COLS.concat([ex]):COLS;
  const linhas=[{cells:cols.map(h=>({v:h[0],t:'s',s:S.hdr})),h:30}]
    .concat(itens.map(i=>linhaAluno(i,plano)));
  const ultima=colName(cols.length-1)+linhas.length;
  const colR=colName(iResolv);
  return {nome:plano.aba,cols:cols.map(h=>h[1]),linhas,congelar:1,
    filtro:'A1:'+ultima,
    validacao:{ref:colR+'2:'+colR+linhas.length,opcoes:['NÃO','SIM'],
      dica:'Marque SIM quando a ação estiver concluída.'},
    formatacao:{ref:colR+'2:'+colR+linhas.length,igual:'SIM'}};
}

function abaResumo(code,itens,usados){
  const L=[];
  const t=(v,s)=>({v,t:'s',s:s||S.txt});
  L.push({cells:[{v:'Plano de Ação · Bolsão 2027',t:'s',s:S.titulo}],h:24});
  L.push({cells:[{v:CODE[code]+' · '+itens.length+' alunos no acompanhamento',t:'s',s:S.legenda}]});
  L.push({cells:[{v:'Fontes cruzadas: Marketplace (pedidos da 1ª cota), planilha TM Bolsão 1 e '+
    'Base de matrículas do Totvs. Rematrículas e cadastros de teste foram excluídos.',t:'s',s:S.legenda}],h:28});
  L.push({cells:[]});
  L.push({cells:[t('Aba',S.hdr),t('Critério',S.hdr),t('Ação necessária',S.hdr),t('Alunos',S.hdr)],h:30});
  PLANOS.concat([SEMPLANO]).forEach(p=>{
    const q=itens.filter(i=>i.plano===p.n).length;
    if(!q&&p.n===0)return;
    L.push({cells:[t(p.aba),t(p.criterio),{v:p.acao,t:'s',s:S.acao},
      {v:q,t:'n',s:S.num}],h:30});
  });
  L.push({cells:[]});
  L.push({cells:[{v:'Como usar',t:'s',s:S.verde}]});
  [ 'Cada aba traz os alunos de um plano, com todos os dados necessários para agir sem abrir outra planilha.',
    'A coluna “Resolvido?” tem lista suspensa: escolha SIM quando a ação estiver concluída — a célula fica verde sozinha.',
    'Preencha também “Quem resolveu” e “Data da conclusão” para o acompanhamento.',
    'A primeira linha está congelada e com filtro: dá para ordenar por dias parado ou por situação no Totvs.',
    '“Dias parado” conta da data do pagamento até '+new Date(HOJE).toLocaleDateString('pt-BR')+'.'
  ].forEach(x=>L.push({cells:[{v:'•  '+x,t:'s',s:S.legenda}]}));
  if(itens.some(i=>i.plano===0)){
    L.push({cells:[]});
    L.push({cells:[{v:'A aba “Sem plano definido” reúne os alunos que não se encaixam em nenhuma das seis '+
      'regras — em geral situações do Totvs que as regras não citam, como “Reserva de vaga - Pré-Mat”. '+
      'Eles estão listados para não sumirem do controle, sem ação atribuída.',t:'s',s:S.legenda}],h:42});
  }
  return {nome:'Resumo',cols:[34,42,54,10],linhas:L,congelar:0};
}

/* ---- geração ---- */
if(!fs.existsSync(SAIDA))fs.mkdirSync(SAIDA,{recursive:true});
const relatorio=[];
for(const code of ORDER){
  const itens=ITEMS.filter(i=>i.code===code);
  if(!itens.length)continue;
  const abas=[abaResumo(code,itens)];
  const contagem={};
  PLANOS.concat([SEMPLANO]).forEach(p=>{
    const grupo=itens.filter(i=>i.plano===p.n)
      .sort((a,b)=>(a.dISO||'9999').localeCompare(b.dISO||'9999')||
        String(a.nome).localeCompare(String(b.nome),'pt'));
    contagem[p.n]=grupo.length;
    if(grupo.length)abas.push(abaPlano(p,grupo));
  });
  const arq=path.join(SAIDA,'Plano de Acao 2027 - '+CODE[code].replace(/[\\\/:*?"<>|]/g,'')+'.xlsx');
  fs.writeFileSync(arq,pasta(abas));
  relatorio.push({code,nome:CODE[code],total:itens.length,abas:abas.length,contagem,
    bytes:fs.statSync(arq).size});
}
const NS=PLANOS.map(p=>p.n).concat([0]);
console.log('unid  total  '+NS.map(n=>(n?'P'+n:'s/p').padStart(4)).join('')+'   abas    KB');
relatorio.forEach(r=>console.log(r.code.padEnd(5),String(r.total).padStart(5),' ',
  NS.map(n=>String(r.contagem[n]||0).padStart(4)).join(String()),
  String(r.abas).padStart(5),String(Math.round(r.bytes/1024)).padStart(6)));
const soma=n=>relatorio.reduce((a,r)=>a+(r.contagem[n]||0),0);
console.log('TOTAL'.padEnd(5),String(relatorio.reduce((a,r)=>a+r.total,0)).padStart(5),' ',
  NS.map(n=>String(soma(n)).padStart(4)).join(String()));
console.log();
PLANOS.concat([SEMPLANO]).forEach(p=>{const q=soma(p.n);
  if(q||p.n)console.log(' '+(p.n?('P'+p.n):'s/plano').padEnd(8),String(q).padStart(4),' ',p.criterio);});
console.log();
console.log('arquivos em:',path.resolve(SAIDA));
