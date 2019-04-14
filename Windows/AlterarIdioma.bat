@echo off  
================================================================================  
:: NOME   : Alterar idioma Windows 10
:: AUTOR  : Ivo Dias  
:: VERSAO : 1.0.MSC 
================================================================================
::Titulo:
title Alterar idioma Windows 10

:: Recebe o hostname
set /p Hostname="Informe o hostname: "

:: Faz a copia do arquivo e cria o log de erro
robocopy  C:\temp\Mudar_Idioma \\%Hostname%\c$\Temp\ "LanguagePack.cab" /R:3 /W:10 /J /V /ETA /TEE /LOG:C:\temp\%Hostname%.log

:: Altera o idioma
psexec \\%Hostname% cmd /c dism /Online /Add-Package /PackagePath:C:\Temp\LanguagePack.cab

::Finalizando
color 27
echo "O script foi executado"
Pause
