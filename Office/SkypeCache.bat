@echo off  
================================================================================  
:: NOME   : Limpar cache do Skype/Lync  
:: AUTOR  : Ivo Dias  
:: VERSAO : 2.0.MSC  
================================================================================  
:: Configurações de tela.  
echo off  
title Limpar cache do Skype  
================================================================================= 
::Verifica as pastas 
if exist "%userprofile%\AppData\Local\Microsoft\Office\15.0\Lync" ( 
	::Office 2013
	cd /d %userprofile%\AppData\Local\Microsoft\Office\15.0\Lync 
	del /s /f /q *.*  
	mkdir Tracing 
) else if exist "%userprofile%\AppData\Local\Microsoft\Office\16.0\Lync" ( 
	::Office 2016
	cd /d %userprofile%\AppData\Local\Microsoft\Office\16.0\Lync  
	del /s /f /q *.*  
	mkdir Tracing 
) else (
	echo "Não foi localizada a pasta"
	goto END
)
ipconfig /flushdns 
:: Encerrando.  
cls  
color 97  
echo.O procedimento foi completado com sucesso.  
echo.Reinicie o computador para completar o procedimento 
:END 
echo.Aperte alguma tecla para encerrar a ferramenta . . .  
pause>nul