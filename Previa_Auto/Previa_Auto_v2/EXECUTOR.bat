@echo off
title Robô de Automação da Prévia FCA

echo.
echo =================================================================
echo.
echo      PASSO UNICO (MANUAL): EXTRACAO DE DADOS DO SAP
echo.
echo =================================================================
echo.
echo   1. Abra o SAP e extraia o relatorio LISTCUBE (YA_CONAN).
echo.
echo   2. Salve o arquivo com o nome EXATO: LISTCUBE_Export.xlsx
echo.
echo      DENTRO DA PASTA: ExtracaoSAP
echo.
echo =================================================================
echo.
echo      APOS SALVAR O ARQUIVO, PRESSIONE QUALQUER TECLA
echo               PARA INICIAR A AUTOMACAO
echo =================================================================
echo.
pause > nul

echo.
echo ================================================
echo      INICIANDO PROCESSAMENTO AUTOMATICO...
echo ================================================
echo.

python automacao_previa.py

echo.
echo ================================================
echo.
echo Processo finalizado. Pressione qualquer tecla para fechar.
echo.
pause > nul