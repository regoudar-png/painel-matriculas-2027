/* Gera o relatório do rastreio de 03/09 a partir do JSON do motor.
 *
 *   node painel/gerar-rastreio.js
 *
 * Nenhum número é digitado à mão: tudo sai de dados/rastreio-0309.json. */

const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const j = JSON.parse(fs.readFileSync(path.join(RAIZ, 'dados', 'rastreio-0309.json'), 'utf8'));
const { CODE, ORDER } = j;

const HOJE = '2026-09-03';
const iso = d => String(d || '').split('/').reverse().join('-');
const brl = v => 'R$ ' + Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const brl0 = v => 'R$ ' + Number(v).toLocaleString('pt-BR', { maximumFractionDigits: 0 });
const nf = v => Number(v).toLocaleString('pt-BR');
const esc = s => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const dm = d => String(d || '').slice(0, 5);

const A = j.dupla.alunos;
const vencDe = a => (a.cota[0] && a.cota[0].venc) || (a.reserva[0] && a.reserva[0].venc) || null;
const vencido = a => iso(vencDe(a)) < HOJE;
const vencidos = A.filter(vencido), aVencer = A.filter(a => !vencido(a));
const soma = l => l.reduce((s, a) => s + a.valorAberto, 0);

/* ---------- agenda de vencimentos ---------- */
const agenda = {};
A.forEach(a => { const d = vencDe(a); (agenda[d] = agenda[d] || { n: 0, v: 0 }).n++; agenda[d].v += a.valorAberto; });
const diasOrd = Object.keys(agenda).sort((x, y) => iso(x) < iso(y) ? -1 : 1);
const picoDia = Math.max(...diasOrd.map(d => agenda[d].v));

/* ---------- por unidade ---------- */
const porUni = {};
A.forEach(a => {
  const o = porUni[a.u] = porUni[a.u] || { n: 0, v: 0, venc: 0, vVenc: 0, nada: 0 };
  o.n++; o.v += a.valorAberto;
  if (vencido(a)) { o.venc++; o.vVenc += a.valorAberto; }
  if (a.aberto === 'as duas') o.nada++;
});
const unidadesOrd = ORDER.filter(u => porUni[u]).sort((x, y) => porUni[y].v - porUni[x].v);
const picoUni = Math.max(...unidadesOrd.map(u => porUni[u].v));

/* ---------- matriz unidade x status ---------- */
const ST = ['Matriculado', 'Reserva de vaga - Matriculado', 'Pré-Matriculado',
  'Reserva de vaga - Pré-Mat', 'Pendente', 'Cancelado'];
const ST_CURTO = { 'Matriculado': 'Matriculado', 'Reserva de vaga - Matriculado': 'RV Matriculado',
  'Pré-Matriculado': 'Pré-matriculado', 'Reserva de vaga - Pré-Mat': 'RV Pré-mat',
  'Pendente': 'Pendente', 'Cancelado': 'Cancelado' };
const FECHADO = new Set(['Matriculado', 'Reserva de vaga - Matriculado']);

const totUni = (bloco, u) => Object.values(bloco.porUnidade[u] || {}).reduce((s, v) => s + v, 0);

function matriz(bloco, rot) {
  const usados = ST.filter(s => ORDER.some(u => (bloco.porUnidade[u] || {})[s]));
  const linhas = ORDER.filter(u => bloco.porUnidade[u]);
  const tot = Object.fromEntries(usados.map(s => [s, 0]));
  let totGeral = 0, totFechado = 0;
  const corpo = linhas.map(u => {
    const r = bloco.porUnidade[u];
    const t = usados.reduce((s, x) => s + (r[x] || 0), 0);
    const fech = usados.filter(s => FECHADO.has(s)).reduce((s, x) => s + (r[x] || 0), 0);
    totGeral += t; totFechado += fech;
    usados.forEach(s => tot[s] += (r[s] || 0));
    const pc = t ? fech / t * 100 : 0;
    return `<tr>
      <th scope="row"><span class="sig">${u}</span>${esc(CODE[u])}</th>
      ${usados.map(s => {
        const v = r[s] || 0;
        return `<td class="${v ? (FECHADO.has(s) ? 'v ok' : s === 'Cancelado' ? 'v out' : 'v') : 'v z'}">${v || '·'}</td>`;
      }).join('')}
      <td class="v tt">${t}</td>
      <td class="prog"><span class="barra"><i style="width:${pc.toFixed(1)}%"></i></span><b>${pc.toFixed(0)}%</b></td>
    </tr>`;
  }).join('');
  const pcT = totGeral ? totFechado / totGeral * 100 : 0;
  return `<div class="tabela-rolagem">
    <table class="mat">
      <caption>${esc(rot)} — ${nf(bloco.n)} alunos com lançamento financeiro</caption>
      <thead><tr><th scope="col">Unidade</th>
        ${usados.map(s => `<th scope="col">${esc(ST_CURTO[s] || s)}</th>`).join('')}
        <th scope="col">Total</th><th scope="col">Fechadas</th></tr></thead>
      <tbody>${corpo}</tbody>
      <tfoot><tr><th scope="row">Rede</th>
        ${usados.map(s => `<td class="v">${tot[s] || '·'}</td>`).join('')}
        <td class="v tt">${totGeral}</td>
        <td class="prog"><span class="barra"><i style="width:${pcT.toFixed(1)}%"></i></span><b>${pcT.toFixed(0)}%</b></td>
      </tr></tfoot>
    </table></div>`;
}

/* ---------- lista nominal ---------- */
const listaHtml = unidadesOrd.map(u => {
  const alunos = A.filter(a => a.u === u).sort((x, y) => {
    const vx = vencido(x), vy = vencido(y);
    return (vy - vx) || (y.valorAberto - x.valorAberto);
  });
  return `<tbody class="grupo-uni">
    <tr class="cab-uni"><th colspan="6"><span class="sig">${u}</span>${esc(CODE[u])}
      <span class="z">${alunos.length} aluno${alunos.length > 1 ? 's' : ''} · ${brl(porUni[u].v)}</span></th></tr>
    ${alunos.map(a => {
      const vc = vencido(a);
      const res = a.reserva[0], cot = a.cota[0];
      return `<tr>
        <td class="al">${esc(a.aluno)}<i>${esc(a.serie)} · ${esc(a.trilha || '')}</i></td>
        <td><span class="pill ${FECHADO.has(a.status) ? 'p-ok' : 'p-warn'}">${esc(ST_CURTO[a.status] || a.status)}</span></td>
        <td class="num ${a.pagoRes > 0 ? 'paga' : 'deve'}">${res ? brl(res.liq) : '—'}<i>${a.dtRes ? 'baixada ' + dm(a.dtRes) : 'não baixada'}</i></td>
        <td class="num deve">${cot ? brl(cot.liq) : '—'}<i>vence ${dm(cot ? cot.venc : '')}</i></td>
        <td>${vc ? '<span class="pill p-crit">vencida</span>' : '<span class="pill p-info">a vencer</span>'}</td>
        <td class="num total">${brl(a.valorAberto)}</td>
      </tr>`;
    }).join('')}
  </tbody>`;
}).join('');

/* ---------- página ---------- */
const html = `<title>Cotas em Aberto na Captação</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;800&display=swap">
<style>
:root{
  --papel:#FBFCFE; --carta:#FFFFFF; --carta-2:#F4F9FD;
  --tinta:#10314F; --tinta-2:#3D6285; --fraco:#7793AC;
  --linha:#D3E2F1; --linha-2:#E8F1F9;
  --azul:#1C61AC; --laranja:#FA4F02; --verde:#009639;
  --tinta-azul:#EDF4FB; --tinta-laranja:#FFF1EA; --tinta-verde:#E7FBF1;
  --sombra:0 1px 2px rgba(16,49,79,.06), 0 8px 24px -16px rgba(16,49,79,.25);
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --papel:#08131F; --carta:#0F2434; --carta-2:#143047;
  --tinta:#DCEAF7; --tinta-2:#A9C6DF; --fraco:#6A8B9E;
  --linha:#1E3E5C; --linha-2:#183349;
  --azul:#5FA8E0; --laranja:#FF8348; --verde:#2FBE76;
  --tinta-azul:#0F2337; --tinta-laranja:#2B1409; --tinta-verde:#0B2418;
  --sombra:0 1px 2px rgba(0,0,0,.4), 0 8px 24px -16px rgba(0,0,0,.7);
}}
:root[data-theme="dark"]{
  --papel:#08131F; --carta:#0F2434; --carta-2:#143047;
  --tinta:#DCEAF7; --tinta-2:#A9C6DF; --fraco:#6A8B9E;
  --linha:#1E3E5C; --linha-2:#183349;
  --azul:#5FA8E0; --laranja:#FF8348; --verde:#2FBE76;
  --tinta-azul:#0F2337; --tinta-laranja:#2B1409; --tinta-verde:#0B2418;
  --sombra:0 1px 2px rgba(0,0,0,.4), 0 8px 24px -16px rgba(0,0,0,.7);
}
*{box-sizing:border-box}
body{margin:0;background:var(--papel);color:var(--tinta);
  font:400 15px/1.6 Poppins,"Century Gothic",-apple-system,Segoe UI,sans-serif;
  -webkit-font-smoothing:antialiased}
.folha{max-width:1080px;margin:0 auto;padding:40px 24px 72px}
h1,h2,h3{font-weight:800;letter-spacing:-.02em;text-wrap:balance;margin:0}
h1{font-size:34px;line-height:1.15;color:var(--azul)}
h2{font-size:21px;line-height:1.25}
h3{font-size:15px;line-height:1.3}
p{margin:0}
.z{color:var(--fraco)}

/* cabeçalho */
.topo{display:flex;flex-wrap:wrap;gap:18px;align-items:flex-end;justify-content:space-between;
  padding-bottom:20px;border-bottom:2px solid var(--azul);margin-bottom:8px}
.topo .sub{max-width:56ch;margin-top:10px;color:var(--tinta-2);font-size:14.5px}
.carimbo{font-size:11.5px;letter-spacing:.08em;text-transform:uppercase;color:var(--fraco);
  text-align:right;line-height:1.9;white-space:nowrap}
.carimbo b{display:block;font-weight:800;color:var(--tinta-2);letter-spacing:.04em}

/* seções */
section{margin-top:44px}
.eyebrow{font-size:11.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--laranja);
  font-weight:800;margin-bottom:6px}
.eyebrow.azul{color:var(--azul)}
.intro{margin-top:10px;color:var(--tinta-2);max-width:68ch;font-size:14.5px}

/* números de topo */
.numeros{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:1px;
  background:var(--linha);border:1px solid var(--linha);border-radius:12px;overflow:hidden;margin-top:20px}
.n{background:var(--carta);padding:16px 18px 17px;display:flex;flex-direction:column;gap:3px}
.n .rot{font-size:10.5px;letter-spacing:.09em;text-transform:uppercase;color:var(--fraco)}
.n .vl{font-size:27px;line-height:1.1;font-weight:800;font-variant-numeric:tabular-nums;
  letter-spacing:-.025em}
.n .pe{font-size:12px;color:var(--tinta-2)}
.n.destaque{background:var(--tinta-laranja)} .n.destaque .vl{color:var(--laranja)}
.n.bom .vl{color:var(--verde)} .n.info .vl{color:var(--azul)}

/* agenda de vencimentos */
.agenda{margin-top:22px;border:1px solid var(--linha);border-radius:12px;overflow:hidden;
  background:var(--carta);box-shadow:var(--sombra)}
.agenda-cab{display:flex;align-items:baseline;gap:10px;padding:14px 18px 4px;flex-wrap:wrap}
.agenda-cab h3{color:var(--azul)}
.agenda-lista{display:flex;flex-direction:column;gap:1px;background:var(--linha-2);
  margin:12px 0 0;border-top:1px solid var(--linha)}
.dia{display:grid;grid-template-columns:96px 1fr 120px 86px;gap:14px;align-items:center;
  background:var(--carta);padding:11px 18px}
.dia.passou{background:var(--tinta-laranja)}
.dia .data{font-weight:800;font-size:13.5px;font-variant-numeric:tabular-nums}
.dia.passou .data{color:var(--laranja)}
.dia .trilho{height:9px;background:var(--linha-2);border-radius:5px;overflow:hidden}
.dia .trilho i{display:block;height:100%;background:var(--azul);border-radius:5px}
.dia.passou .trilho i{background:var(--laranja)}
.dia .din{text-align:right;font-variant-numeric:tabular-nums;font-size:13.5px;font-weight:800}
.dia .qtd{text-align:right;font-size:12.5px;color:var(--tinta-2);font-variant-numeric:tabular-nums}

/* tabelas */
.tabela-rolagem{overflow-x:auto;margin-top:20px;border:1px solid var(--linha);border-radius:12px;
  background:var(--carta);box-shadow:var(--sombra)}
table{border-collapse:collapse;width:100%;font-size:13.5px}
caption{caption-side:top;text-align:left;padding:14px 18px 12px;font-weight:800;font-size:14px;
  color:var(--azul);border-bottom:1px solid var(--linha)}
thead th{background:var(--azul);color:#fff;font-weight:800;font-size:10.5px;letter-spacing:.07em;
  text-transform:uppercase;padding:10px 12px;text-align:right;white-space:nowrap}
thead th:first-child{text-align:left}
tbody th,tfoot th{text-align:left;font-weight:400;padding:9px 12px;white-space:nowrap}
td{padding:9px 12px;text-align:right;font-variant-numeric:tabular-nums;border-top:1px solid var(--linha-2)}
tbody tr:nth-child(even) td,tbody tr:nth-child(even) th{background:var(--tinta-azul)}
tfoot th,tfoot td{background:var(--tinta-verde);border-top:2px solid var(--verde);font-weight:800}
.sig{display:inline-block;min-width:26px;padding:1px 5px;margin-right:8px;border-radius:4px;
  background:var(--tinta-azul);color:var(--azul);font-weight:800;font-size:10.5px;
  letter-spacing:.04em;text-align:center;vertical-align:1px}
.v.z{color:var(--fraco)} .v.ok{color:var(--verde);font-weight:800}
.v.out{color:var(--laranja)} .v.tt{font-weight:800}
.prog{display:flex;align-items:center;gap:8px;justify-content:flex-end;min-width:120px}
.prog .barra{flex:1;height:7px;background:var(--linha-2);border-radius:4px;overflow:hidden;min-width:52px}
.prog .barra i{display:block;height:100%;background:var(--verde);border-radius:4px}
.prog b{font-weight:800;font-size:12.5px;min-width:34px;text-align:right}

/* dinheiro por unidade */
.uni{display:flex;flex-direction:column;gap:1px;background:var(--linha-2);margin-top:20px;
  border:1px solid var(--linha);border-radius:12px;overflow:hidden;box-shadow:var(--sombra)}
.uni-cab,.uni-l{display:grid;grid-template-columns:190px 1fr 96px 118px 118px;gap:14px;
  align-items:center;background:var(--carta);padding:11px 18px}
.uni-cab{background:var(--tinta-azul);font-size:10.5px;letter-spacing:.07em;text-transform:uppercase;
  color:var(--azul);font-weight:800}
.uni-cab span:not(:first-child),.uni-l>*:not(:first-child):not(.trilho2){text-align:right}
.uni-l .nome{font-size:13.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.trilho2{height:9px;background:var(--linha-2);border-radius:5px;overflow:hidden;display:flex}
.trilho2 i{display:block;height:100%}
.trilho2 .vc{background:var(--laranja)} .trilho2 .av{background:var(--azul)}
.uni-l .din{font-variant-numeric:tabular-nums;font-weight:800;font-size:13.5px}
.uni-l .sec{font-variant-numeric:tabular-nums;font-size:12.5px;color:var(--tinta-2)}
.uni-l .sec.vencida{color:var(--laranja);font-weight:800}
.uni-rodape{display:grid;grid-template-columns:190px 1fr 96px 118px 118px;gap:14px;
  background:var(--tinta-verde);border-top:2px solid var(--verde);padding:12px 18px;font-weight:800}
.uni-rodape>*:not(:first-child){text-align:right;font-variant-numeric:tabular-nums}

/* lista nominal */
.nominal{margin-top:22px;border:1px solid var(--linha);border-radius:12px;overflow:hidden;
  background:var(--carta);box-shadow:var(--sombra)}
.nominal .rolagem{max-height:620px;overflow:auto}
.nominal table{font-size:13px}
.nominal thead th{position:sticky;top:0;z-index:2}
.cab-uni th{background:var(--tinta-azul);color:var(--azul);font-weight:800;font-size:12.5px;
  padding:9px 12px;border-top:1px solid var(--linha)}
.cab-uni .z{font-weight:400;margin-left:8px;font-size:12px}
.nominal tbody tr:nth-child(even) td{background:transparent}
.nominal tbody tr:not(.cab-uni):hover td{background:var(--carta-2)}
td.al{text-align:left;line-height:1.35}
td.al i{display:block;font-style:normal;font-size:11.5px;color:var(--fraco);margin-top:1px}
td.num i{display:block;font-style:normal;font-size:11px;color:var(--fraco);font-weight:400;margin-top:1px}
td.num{font-weight:800}
td.paga{color:var(--verde)} td.deve{color:var(--laranja)} td.total{color:var(--tinta)}
.pill{display:inline-block;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:800;
  white-space:nowrap;line-height:1.5}
.p-ok{background:var(--tinta-verde);color:var(--verde)}
.p-warn{background:var(--tinta-azul);color:var(--azul)}
.p-crit{background:var(--tinta-laranja);color:var(--laranja)}
.p-info{background:var(--tinta-azul);color:var(--tinta-2)}

/* caixas de texto */
.nota{margin-top:22px;background:var(--carta);border:1px solid var(--linha);border-left:3px solid var(--azul);
  border-radius:0 10px 10px 0;padding:16px 20px;font-size:13.5px;line-height:1.65;color:var(--tinta-2);
  max-width:78ch}
.nota.alerta{border-left-color:var(--laranja);background:var(--tinta-laranja)}
.nota h3{color:var(--azul);margin-bottom:6px;font-size:14px}
.nota.alerta h3{color:var(--laranja)}
.nota ul{margin:8px 0 0;padding-left:18px} .nota li{margin-top:4px}
.nota b{font-weight:800;color:var(--tinta)}

.rodape{margin-top:52px;padding-top:18px;border-top:1px solid var(--linha);
  font-size:12.5px;color:var(--fraco);line-height:1.7;max-width:78ch}
.rodape b{color:var(--tinta-2);font-weight:800}

@media (max-width:820px){
  .folha{padding:28px 16px 56px}
  h1{font-size:27px}
  .dia{grid-template-columns:80px 1fr 96px;gap:10px} .dia .qtd{display:none}
  .uni-cab,.uni-l,.uni-rodape{grid-template-columns:140px 1fr 104px}
  .uni-cab span:nth-child(3),.uni-l .sec,.uni-rodape>*:nth-child(3),
  .uni-cab span:nth-child(5),.uni-rodape>*:nth-child(5){display:none}
}
@media print{body{background:#fff}.nominal .rolagem{max-height:none}}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style>

<div class="folha">

  <header class="topo">
    <div>
      <h1>Cotas em aberto na captação</h1>
      <p class="sub">Quem pagou a reserva de vaga do bolsão e recebeu, depois, uma segunda
        cobrança de entrada que ainda não foi quitada — e como está a matrícula 2027 no resto
        da base financeira.</p>
    </div>
    <div class="carimbo">
      <b>${esc(j.fontes.base)}</b>
      baixas conferidas em ${esc(j.fontes.baixas)}<br>
      posição de 3 de setembro de 2026
    </div>
  </header>

  <section>
    <p class="eyebrow">1 · A dupla cobrança</p>
    <h2>${A.length} alunos com duas cobranças de entrada</h2>
    <p class="intro">Pagaram a reserva de vaga no bolsão, por R$ 300, e o contrato gerou uma
      <b>1ª cota de mensalidade</b> por cima. É essa segunda cobrança que está em aberto —
      valor cheio de mensalidade, não os R$ 300 da reserva.</p>

    <div class="numeros">
      <div class="n destaque">
        <span class="rot">Em aberto</span>
        <span class="vl">${brl0(j.dupla.valorAberto)}</span>
        <span class="pe">${A.length} alunos · ${brl(j.dupla.valorAberto)}</span>
      </div>
      <div class="n destaque">
        <span class="rot">Já vencido</span>
        <span class="vl">${brl0(soma(vencidos))}</span>
        <span class="pe">${vencidos.length} alunos · venceu em ${dm(diasOrd[0])}</span>
      </div>
      <div class="n info">
        <span class="rot">A vencer</span>
        <span class="vl">${brl0(soma(aVencer))}</span>
        <span class="pe">${aVencer.length} alunos · até ${dm(diasOrd[diasOrd.length - 1])}</span>
      </div>
      <div class="n bom">
        <span class="rot">Reserva já paga</span>
        <span class="vl">${A.filter(a => a.pagoRes > 0).length}</span>
        <span class="pe">de ${A.length} — o dinheiro do bolsão entrou</span>
      </div>
    </div>

    <div class="nota alerta">
      <h3>${A.filter(a => a.aberto === 'as duas').length} alunos não pagaram nem a reserva nem a cota</h3>
      São os mais urgentes: entraram no bolsão, foram lançados no Totvs, e nenhuma das duas
      cobranças foi quitada.
      <ul>${A.filter(a => a.aberto === 'as duas').map(a =>
        `<li><b>${esc(a.aluno)}</b> — ${esc(CODE[a.u])} · ${esc(a.serie)} · ${esc(ST_CURTO[a.status] || a.status)} · ${brl(a.valorAberto)}</li>`).join('')}</ul>
    </div>

    <div class="agenda">
      <div class="agenda-cab">
        <h3>Quando vence</h3>
        <span class="z">${vencidos.length} de ${A.length} já passaram do vencimento</span>
      </div>
      <div class="agenda-lista">
        ${diasOrd.map(d => {
          const o = agenda[d], passou = iso(d) < HOJE;
          return `<div class="dia${passou ? ' passou' : ''}">
            <span class="data">${esc(dm(d))}</span>
            <span class="trilho"><i style="width:${(o.v / picoDia * 100).toFixed(1)}%"></i></span>
            <span class="din">${brl(o.v)}</span>
            <span class="qtd">${o.n} aluno${o.n > 1 ? 's' : ''}</span>
          </div>`;
        }).join('')}
      </div>
    </div>

    <div class="uni">
      <div class="uni-cab"><span>Unidade</span><span>vencido / a vencer</span><span>Alunos</span>
        <span>Vencido</span><span>Em aberto</span></div>
      ${unidadesOrd.map(u => {
        const o = porUni[u];
        return `<div class="uni-l">
          <span class="nome"><span class="sig">${u}</span>${esc(CODE[u])}</span>
          <span class="trilho2">
            <i class="vc" style="width:${(o.vVenc / picoUni * 100).toFixed(1)}%"></i>
            <i class="av" style="width:${((o.v - o.vVenc) / picoUni * 100).toFixed(1)}%"></i>
          </span>
          <span class="sec">${o.n}</span>
          <span class="sec${o.vVenc ? ' vencida' : ''}">${o.vVenc ? brl(o.vVenc) : '—'}</span>
          <span class="din">${brl(o.v)}</span>
        </div>`;
      }).join('')}
      <div class="uni-rodape"><span>Rede</span><span></span><span>${A.length}</span>
        <span>${brl(soma(vencidos))}</span><span>${brl(j.dupla.valorAberto)}</span></div>
    </div>

    <div class="nominal">
      <div class="agenda-cab"><h3>Os ${A.length} alunos, um a um</h3>
        <span class="z">por unidade, vencidos primeiro</span></div>
      <div class="rolagem">
        <table>
          <thead><tr>
            <th scope="col">Aluno</th><th scope="col">Situação no Totvs</th>
            <th scope="col">Reserva</th><th scope="col">1ª cota</th>
            <th scope="col">Prazo</th><th scope="col">Em aberto</th>
          </tr></thead>
          ${listaHtml}
        </table>
      </div>
    </div>
  </section>

  <section>
    <p class="eyebrow azul">2 · A matrícula por unidade e status</p>
    <h2>${nf(j.total)} alunos com lançamento financeiro em 2027</h2>
    <p class="intro">A coluna <b>Fechadas</b> é a fatia já em Matriculado ou Reserva de vaga -
      Matriculado. Pré-matriculado tem contrato assinado e falta fechar; Pendente é lançamento
      aberto e não concluído.</p>

    ${matriz(j.cap, 'Captação')}
    ${matriz(j.rem, 'Rematrícula')}

    <div class="nota">
      <h3>O que salta nesta base</h3>
      <ul>
        <li><b>Américas fechou tudo:</b> ${totUni(j.cap, 'AM')} alunos, todos em
          Matriculado, sem nenhum pendente — e é a única unidade sem dupla cobrança.</li>
        <li><b>A rematrícula está meio a meio:</b> ${j.rem.porStatus['Matriculado'] || 0} matriculados
          contra ${j.rem.porStatus['Pré-Matriculado'] || 0} pré-matriculados. Metade tem contrato
          assinado e ainda não fechou.</li>
        <li><b>Bangu e Campo Grande concentram a dupla cobrança:</b>
          ${porUni.BG.n + porUni.CG.n} dos ${A.length} alunos e
          ${brl(porUni.BG.v + porUni.CG.v)} dos ${brl(j.dupla.valorAberto)} em aberto.</li>
      </ul>
    </div>
  </section>

  <p class="rodape">
    <b>De onde vêm os números.</b> A base de 3 de setembro traz o lançamento financeiro de cada
    aluno, mas não diz o que foi pago — quem diz é a ficha financeira de 2 de setembro, cruzada
    por RA. Uma cobrança conta como baixada quando tem valor baixado maior que zero e data de
    baixa válida; parcelas repetidas no export, com o mesmo identificador e dois status de
    boleto, contam uma vez só. Cadastros de teste ficam de fora.<br><br>
    <b>O que esta base não cobre.</b> Só entram alunos com lançamento financeiro em 2027 —
    ${nf(j.total)} pessoas. Quem foi matriculado e ainda não teve contrato gerado não aparece
    aqui, então estes totais são menores que os do painel de acompanhamento.
  </p>

</div>
`;

const destino = path.join(RAIZ, 'relatorio-rastreio-0309.html');
fs.writeFileSync(destino, html);
console.log('->', destino, Math.round(html.length / 1024) + ' KB');
console.log('dupla cobrança:', A.length, 'alunos ·', brl(j.dupla.valorAberto),
  '| vencido:', vencidos.length, brl(soma(vencidos)));
