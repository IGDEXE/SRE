﻿# Shopping GIAM
# AUTOR  : Ivo Dias 
# VERSAO : 01.08.18.MSC 

# Documentação para o LOG
$hostname = Get-Content C:\Fiscal\Config\Hostname.SID # Receive the hostname or IP
$Installpath = Get-Content C:\Fiscal\Config\user.SID # Receive the path
$nome = "GIAM"
$fullname = "GIAM TO Eletronica"
$versao = "10 - 10.10.18"
$logPath = "$Installpath\LOG\$hostname.$nome.log"
$fullpath = "$Installpath\Install\$nome"
$instalador = "instalador.exe"
$date = Get-Date

# Inicia o Log
Add-Content -Path $logPath -Value "|*********************** $fullname ***********************|"
Add-Content -Path $logPath -Value "Start: $date"
Add-Content -Path $logPath -Value "Software: $nome"
Add-Content -Path $logPath -Value "Version: $versao"

# Envia o arquivo
robocopy "$fullpath" \\$hostname\c$\Temp\ "$instalador" /R:3 /W:10 /J /V /ETA /TEE /LOG+:$logPath

# Inicia a instalação
$process = psexec \\$hostname -s cmd /c C:\temp\$instalador /verysilent /norestart 
Add-Content -Path $logPath -Value "Return: $process"

# Finaliza o Log
$date = Get-Date
Add-Content -Path $logPath -Value "Complete: $date"

# Informa o analista que usou o script
$myhostname = get-content env:computername
$msg = "A instalação do $nome, no computador $hostname, foi concluida, mais detalhes disponiveis em $logPath"
Invoke-WmiMethod ` -Path Win32_Process ` -Name Create ` -ArgumentList "msg * $msg" ` -ComputerName $myhostname


