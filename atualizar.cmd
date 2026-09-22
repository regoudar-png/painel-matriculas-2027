@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
rem  Atualiza o painel do dia e guarda o codigo no GitHub.
rem
rem    atualizar.cmd "22 de setembro"
rem
rem  Sem argumento, usa a pasta mais recente de "Matriculas 2027".
rem  Os tres relatorios sao reconhecidos pelo conteudo, nao pelo nome:
rem  cadastro (99 colunas), ficha financeira (49 colunas, CODCOLIGADA em A1)
rem  e o itens_da_venda da Layers. O relatorio de bolsas, se estiver na pasta,
rem  e reconhecido e ignorado.

cd /d "%~dp0"
set "BASE=%USERPROFILE%\Desktop\Matrículas 2027"

if "%~1"=="" (
  for /f "delims=" %%d in ('dir /b /ad /o-d "%BASE%" 2^>nul') do (
    set "PASTA=%BASE%\%%d"
    goto :achou
  )
  echo Nao encontrei nenhuma pasta em "%BASE%".
  exit /b 1
) else (
  set "PASTA=%BASE%\%~1"
)
:achou

if not exist "%PASTA%" (
  echo Pasta nao encontrada: %PASTA%
  exit /b 1
)

echo.
echo === Relatorios de: %PASTA%
echo.

rem  Passa todos os .xlsx da pasta; o motor escolhe o que serve.
set "ARQS="
for %%f in ("%PASTA%\*.xlsx" "%PASTA%\*.XLSX") do set "ARQS=!ARQS! "%%f""

node motor\montar-painel.js !ARQS!
if errorlevel 1 (
  echo.
  echo Falhou ao montar o painel. Nada foi publicado.
  exit /b 1
)

echo.
node painel\montar-app.js dados\painel.json
if errorlevel 1 (
  echo.
  echo Falhou ao montar a pagina. Nada foi publicado.
  exit /b 1
)

echo.
echo === Guardando o codigo no GitHub
git add -A
git diff --cached --quiet && (
  echo    codigo sem alteracao desde o ultimo envio
) || (
  for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set "HOJE=%%a/%%b/%%c"
  git commit -q -m "Painel de !HOJE!" -m "Relatorios de %~1"
  git push -q origin main && echo    enviado || echo    nao consegui enviar; o commit ficou local
)

echo.
echo === Pronto
echo    painel-app.html montado. Falta publicar no link:
echo    abra o painel no claude.ai, clique em Importar relatorios,
echo    solte os tres arquivos da pasta e publique.
echo.
pause
