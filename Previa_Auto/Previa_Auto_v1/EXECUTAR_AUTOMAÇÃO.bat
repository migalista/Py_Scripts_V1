@echo off
title Robô de Automação Assistida da Prévia

:: --- INÍCIO - LÓGICA DE AGENDAMENTO ---
echo.
echo ================================================
echo           VERIFICANDO AGENDAMENTO
echo ================================================
setlocal
set day_str=%date:~0,2%
set /a day_num=1%day_str% - 100
set run_today=0
echo Data de hoje: %date% (Dia do mes: %day_num%)
if %day_num% LEQ 15 (
    echo Primeira quinzena do mes.
    set /a is_odd=(%day_num% %% 2)
    if %is_odd% NEQ 0 (
        echo Dia impar. [PERMITIDO EXECUTAR]
        set run_today=1
    ) else (
        echo Dia par. [NAO EXECUTAR]
    )
) else (
    echo Segunda quinzena do mes. [PERMITIDO EXECUTAR]
    set run_today=1
)
if %run_today% EQU 0 (
    echo.
    echo Processo nao sera executado hoje conforme agendamento.
    goto end_process
)
endlocal
:: --- FIM - LÓGICA DE AGENDAMENTO ---