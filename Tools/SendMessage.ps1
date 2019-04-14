# Script para enviar mensagens  
# AUTOR  : Ivo Dias 
# VERSAO : 1.2.MSC 

# Informe a mensagem
write-host "Bem vindo"
$msg = read-host "Informe sua mensagem: "
$perfil = read-host "Informe o seu nome de usuario: "

# Limpa a tela e exibe as opções de contatos
cls 
write-host "Lista de contatos 1"
write-host "Lista de contatos 2"
write-host "Lista de contatos 3"
$indice = read-host  "Informe o número da lista: "

# De acordo com a opção escolhida, a lista de contatos é definida
if ($indice -match "1")
{ 
     $computers = Get-Content -Path C:\Users\$perfil\Desktop\lista.txt
 
}
if ($indice -match "2")
{ 
     $computers = Get-Content -Path C:\Users\$perfil\Desktop\lista2.txt
}
if ($indice -match "3")
{ 
     $computers = Get-Content -Path C:\Users\$perfil\Desktop\lista3.txt
}
  else
  { 
     Write-Host "Escolha um valor valido"
     exit
  }

# Envia a mensagem para os Hostnames presentes na listagem
try 
{
   ForEach ($computer in $computers) { Invoke-WmiMethod ` -Path Win32_Process ` -Name Create ` -ArgumentList "msg * 
$msg" ` -ComputerName $computer }
} catch  {
   Write-Host A mensagem não foi enviada para $computer 
   pause
}
