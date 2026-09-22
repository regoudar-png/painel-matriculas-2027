# Painel de Matrículas — motor e artefatos

Código que faz a conferência das matrículas (captação + renovação) cruzando três bases.
Salvo aqui em 26/08/2026, tirado da pasta temporária da sessão onde foi escrito.

## Pastas

- `motor/` — o código que faz o trabalho
  - `ler-xlsx.js` — leitor de .xlsx sem dependências (descompacta e lê o XML)
  - `escrever-xlsx.js` — escritor de .xlsx sem dependências (gera as planilhas de plano de ação)
  - `cruzar.js` — cruzamento das três bases, pareamento por nome e categorização
  - `gerar-planos.js` — gera uma planilha de plano de ação por unidade
- `painel/` — o painel web
  - `painel.tpl.html` — template; o marcador `__DATA__` é trocado pelo JSON
  - `painel-atual.html` — versão publicada em 26/08/2026, autônoma
- `dados/`
  - `snapshot-2026-08-24.json` — dados derivados do cruzamento daquela data

## As três fontes

1. **Marketplace (Layers)** — pedidos da 1ª cota. O usuário chama de Marketplace, não de gateway.
2. **Planilha das secretarias** (`TM BOLSÃO 1.xlsx`) — uma aba por unidade, coluna FLUXO.
3. **Relatório do Totvs** (`Base de matrículas.XLSX`) — captação e renovação.

## Limpeza obrigatória do Totvs

Sem isso qualquer contagem sai errada:

- Excluir linhas com `TIPO MATRICULA = REMATRÍCULA` **e** `SITUACAO = Pendente` para a
  conferência de **captação** — são alunos de 2026 apenas liberados, sem movimentação.
  Para a visão de **renovação**, essas linhas são justamente o universo elegível.
- Deduplicar por RA mantendo a situação mais avançada.
- Excluir cadastros de teste (nome contém "TESTE") — eram 91 em 24/08/2026, todos cancelados.

## Siglas das unidades

BG Bangu · CG Campo Grande · CX Duque de Caxias · MD Madureira · NI Nova Iguaçu ·
RM Rocha Miranda · RT Retiro dos Artistas · SJ São João de Meriti · TQ Taquara ·
TJ Tijuca · AM Américas

Duas pegadinhas: **Américas aparece no Totvs como "Colégio e Curso Matriz Educação - Recreio"**
e não tem aba na planilha das secretarias. São João de Meriti não vende pelo Marketplace.

## Situações do Totvs

Contam como matrícula efetivada: `Matriculado`, `Pré-Matriculado`,
`Reserva de vaga - Matriculado`, `Reserva de vaga - Pré-Mat`.

Semântica que sustenta os planos de ação:
- `Pré-Matriculado` e `Reserva de vaga - Pré-Mat` = contrato assinado
  (pagou → baixar a cota; não pagou → cobrar)
- `Pendente` e `Reserva de vaga - Pendente` = contrato ainda não assinado

## Turmas — pendência aberta

Um aluno aparece em várias linhas quando ainda está alocado em turmas candidatas da mesma
série: 3.033 dos 3.034 casos de linha repetida variam só a TURMA. A turma se resolve quando a
matrícula avança — 142 de 143 matrículas efetivadas têm turma única, contra 38% das
rematrículas pendentes.

Regra de contagem: o aluno conta **uma vez**, como elegível na série; a turma só é atribuída
quando ela for única.

O **nome** da turma (3ª série militar, vestibular, etc.) deveria vir na coluna `GRADE`, mas no
export de 24/08/2026 ela estava preenchida em apenas 2 de 10.227 linhas. É preciso um novo
export do Totvs com esse campo populado.

## Pendências do motor

- `cruzar.js` ainda depende de pastas fixas (`x1`, `x2`, `x4`) com os .xlsx já descompactados.
  Precisa aceitar caminhos de arquivo como argumento.
- A visão de renovação ainda não existe: hoje as rematrículas são descartadas.
