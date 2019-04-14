@echo off  
:: *********************************************************************************************
:: NOME   : Redefinindo modern authentication  
:: AUTOR  : Ivo Dias  
:: VERSAO : 1.0.MSC  
:: *********************************************************************************************
:Menu
echo.     1. Desativar ADAL Office 2013 
echo.     2. Ativar ADAL Office 2013 
echo.     3. Desativar ADAL Office 2016
echo.     4. Ativar ADAL Office 2016
echo. 
 
set /p option=Escolha uma opcao:  
 
if %option% EQU 1 ( 
    call :OUT2013 
) else if %option% EQU 2 ( 
    call :INI2013 
) else if %option% EQU 3 ( 
    goto :OUT2016 
) else if %option% EQU 4 ( 
    call :INI2016 
) else ( 
    echo. 
    echo.Opcao invalida. 
    echo. 
    echo.Aperte alguma tecla para continuar . . . 
    pause>nul 
) 

goto Menu

:: *********************************************************************************************

:OUT2013
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\15.0\Common\Identity\EnableADAL" /f /v EnumDevice1 /t REG_DWORD /d 0 
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\15.0\Common\Identity\Version" /f /v EnumDevice1 /t REG_DWORD /d 0
reg.exe add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\OneDrive" /f /v EnableADAL /t REG_DWORD /d 0
echo O procedimento foi completado com sucesso
pause
goto :FIM

:INI2013
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\15.0\Common\Identity\EnableADAL" /f /v EnumDevice1 /t REG_DWORD /d 1 
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\15.0\Common\Identity\Version" /f /v EnumDevice1 /t REG_DWORD /d 1
reg.exe add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\OneDrive" /f /v EnableADAL /t REG_DWORD /d 1
echo O procedimento foi completado com sucesso
pause
goto :FIM

:OUT2016
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Identity\EnableADAL" /f /v EnumDevice1 /t REG_DWORD /d 0 
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Identity\Version" /f /v EnumDevice1 /t REG_DWORD /d 0
reg.exe add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\OneDrive" /f /v EnableADAL /t REG_DWORD /d 0
echo O procedimento foi completado com sucesso
pause
goto :FIM

:INI2016
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Identity\EnableADAL" /f /v EnumDevice1 /t REG_DWORD /d 1 
reg.exe add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Identity\Version" /f /v EnumDevice1 /t REG_DWORD /d 1
reg.exe add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\OneDrive" /f /v EnableADAL /t REG_DWORD /d 1
echo O procedimento foi completado com sucesso
pause
goto :FIM

:FIM
exit