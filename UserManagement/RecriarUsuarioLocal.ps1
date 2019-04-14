# Script para remoção de usuarios 
# AUTOR  : Ivo Dias
# VERSAO : 1.7.MSC
# IDIOMA : PT-BR

#Limpa a tela antes de iniciar o programa
cls

#Recebe o nome do usuario e armazena em uma variavel
$objUser = New-Object System.Security.Principal.NTAccount(Read-Host -Prompt "Enter Username")

#Entra na pasta de usuarios 
cd C:\users
try 
{
    #Entra na pasta de usuarios e renomeia a pasta do User informado para .old
    Rename-Item -Path "$($objUser.value)" -NewName "$($objUser.value).old"
    #Identifica o SID do usuario e remove ele
    $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
    Remove-Item -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\$($strSID.Value)"
    #Escreve na tela informando que encerrou
    Write-Host O procedimento foi completado com sucesso
} 

catch  
{
   #Em caso de erro, uma mensagem é exibida para o usuario. 
   write-host "O perfil não foi localizado, verifique se digitou corretamente o nome do usuario"
}

Pause
