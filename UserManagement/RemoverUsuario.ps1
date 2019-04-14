# Script para remoção de pastas de usuarios 
# AUTOR  : Ivo Dias
# VERSAO : 2.1.MSC

#Limpa a tela antes de iniciar o programa
cls

#Pega o nome do computador em uma lista
#Essa lista precisa ter o nome "lista.txt" e estar armazenada na Temp
$computerlist = Get-Content C:\Temp\lista.txt

#Recebe o nome do usuario e armazena em uma variavel
$objUser = New-Object System.Security.Principal.NTAccount(Read-Host -Prompt "Informe o usuario: ")

#Limpa a tela e exibe uma mensagem enquanto carrega
cls
Write-Host "Aguarde a conclusão do procedimento, não feche está tela"
#Salva data e hora do inicio da aplicação
$data = Get-Date
#Grava em um Log o momento em que a aplicação foi iniciada
Add-Content -Path C:\Temp\LogRemFolder.txt -Value "|*********************************************|" 
Add-Content -Path C:\Temp\LogRemFolder.txt -Value "A ferramenta de remoção iniciou em $data"

#Localiza o SID do usuario por meio do nome
try
{
    #Converte o nome do usuario em um SID
    $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
}
catch
{
    #Verifica a data
    $data = Get-Date
    #Caso um erro ocorra, ele registra no Log
    Add-Content -Path C:\Temp\LogRemFolder.txt -Value "Não foi possivel localizar o SID do usuario $objUser em $data."
    Add-Content -Path C:\Temp\LogRemFolder.txt -Value "<<<<<<<<<<<<<<<   Erro   >>>>>>>>>>>>>>>"
    Write-Host "O usuario $objUser não é valido, mais detalhes em C:\Temp\LogRemFolder.txt"
    Pause
    Exit
}
#Na listagem de computadores: 
foreach ($computer in $computerlist){

    #Verifica a data
    $data = Get-Date
    try 
    {
        #Remove a pasta no computador
        Get-WmiObject -ComputerName $computer win32_userprofile| Where-Object {$_.SID -eq $($strSID.Value)} | ForEach {$_.Delete()} # -ErrorAction Stop
        #Salva no Log as informações sobre a exclusão 
        Add-Content -Path C:\Temp\LogRemFolder.txt -Value "Usuario: $($objUser.Value) |Computador: $computer | Removido em $data"
        #Mostra uma mensagem de confirmação na tela
        Write-Host "O usuario $($objUser.Value) foi removido com sucesso do computador $computer." 
    }
    catch
    {
        #Salva o erro
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        #Armazena a informação no Log
        Add-Content -Path C:\Temp\LogRemFolder.txt -Value "Usuario: $($objUser.Value) | Computador: $computer | Erro ao remover $data"
        Add-Content -Path C:\Temp\LogRemFolder.txt -Value "Mensagem de erro: $ErrorMessage | Item: $FailedItem"
        Add-Content -Path C:\Temp\LogRemFolder.txt -Value "<<<<<<<<<<<<<<<   Erro   >>>>>>>>>>>>>>>"
        #Informa ao usuario para verificar o Log
        Write-Host "Um erro ocorre ao remover o $($objUser.Value) no computador $computer"
        Write-Host "Mais detalhes em C:\Temp\LogRemFolder.txt"
    }
 }

#Verifica a data
$data = Get-Date
#Grava em um Log o momento em que a aplicação foi encerrada
Add-Content -Path C:\Temp\LogRemFolder.txt -Value "A ferramenta de remoção encerrou em $data" 
Add-Content -Path C:\Temp\LogRemFolder.txt -Value "!=================================================!"
Pause