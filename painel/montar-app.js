/* Monta o painel publicável: layout + motor no navegador + tela de importação.
 *
 *   node painel/montar-app.js [dados/painel.json]
 *
 * Gera dois arquivos:
 *   painel-app.html      fragmento, é o que vai para o Artifact na 1ª publicação
 *   painel-app.doc.html  documento completo, só para conferir no navegador
 *
 * A página carrega dentro de si o próprio molde (em base64, no script #fonte).
 * Quando alguém importa relatórios novos, ela reescreve esse molde com os dados
 * novos e publica — e a versão publicada continua sabendo se republicar. */

const fs = require('fs');
const path = require('path');

const AQUI = __dirname;
const RAIZ = path.join(AQUI, '..');
const ler = p => fs.readFileSync(p, 'utf8');

const dados = process.argv[2] || path.join(RAIZ, 'dados', 'painel.json');
const json = ler(dados).replace(/<\//g, '<\\u002f');

// Reempacota o motor sempre. Já aconteceu de eu mexer no cruzar.js e publicar
// a página com o motor antigo dentro: os dados vinham do motor novo e a
// importação da própria página, do velho.
require('child_process').execFileSync(process.execPath,
  [path.join(RAIZ, 'motor', 'empacotar-web.js')], { stdio: 'inherit' });

const layout = ler(path.join(AQUI, 'layout-artifact.tpl.html'));
const motor = ler(path.join(AQUI, 'motor-web.js'));
const importador = ler(path.join(AQUI, 'importar.html'));

/* ---- fragmento: layout + motor + importador + molde ---- */
// O importador precisa de D (do script do layout) e de MotorPainel: entra por
// último. O molde fica num script text/plain, que o navegador não interpreta.
const fragmento = layout.trimEnd() + '\n\n' +
  '<script>\n' + motor + '</script>\n\n' +
  importador.trimEnd() + '\n\n' +
  '<script id="fonte" type="text/plain">__TPL__</script>\n';

/* ---- documento completo, o formato que artifact.publish() exige ---- */
const documento =
`<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Painel de Matrículas 2027</title>
<style>:root{color-scheme:light dark}body{margin:0;font:14px system-ui,-apple-system,"Segoe UI",sans-serif}
img{max-width:100%}[hidden]{display:none!important}</style>
</head>
<body>
${fragmento}</body>
</html>
`;

// Substituição literal: o replacement vai por função para que $& e $' dentro
// do JSON não sejam interpretados como referência de captura.
const põe = (s, marca, valor) => s.replace(marca, () => valor);
const b64 = s => Buffer.from(s, 'utf8').toString('base64');

const moldeB64 = b64(documento);
const saidaFrag = põe(põe(fragmento, '__DATA__', json), '__TPL__', moldeB64);
const saidaDoc = põe(põe(documento, '__DATA__', json), '__TPL__', moldeB64);

const a = path.join(RAIZ, 'painel-app.html');
const b = path.join(RAIZ, 'painel-app.doc.html');
fs.writeFileSync(a, saidaFrag);
fs.writeFileSync(b, saidaDoc);

const kb = s => Math.round(s.length / 1024) + ' KB';
console.log('dados   ', path.basename(dados), kb(json));
console.log('motor   ', kb(motor));
console.log('molde   ', kb(documento), '-> base64', kb(moldeB64));
console.log('->', a, kb(saidaFrag));
console.log('->', b, kb(saidaDoc));
if (saidaDoc.length > 15.5 * 1024 * 1024) console.log('ATENÇÃO: perto do limite de 16 MB do Artifact');
