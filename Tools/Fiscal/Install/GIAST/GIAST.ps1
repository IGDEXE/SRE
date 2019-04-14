# Shopping GIAM
# AUTOR  : Ivo Dias 
# VERSAO : 20.08.18.MSC 

# Documentação para o LOG
$hostname = Get-Content C:\Fiscal\Config\Hostname.SID # Receive the hostname or IP
$Installpath = Get-Content C:\Fiscal\Config\user.SID # Receive the path
$nome = "GIAST"
$fullname = "GIA ST"
$versao = "3"
$logPath = "$Installpath\LOG\$hostname.$nome.log"
$fullpath = "$Installpath\Install\$nome"
$instalador = "InstalaGIAST3.exe"
$pasta = "C:\Program Files (x86)\GIAST3"
$date = Get-Date

# Inicia o Log
Add-Content -Path $logPath -Value "|*********************** $fullname ***********************|"
Add-Content -Path $logPath -Value "Start: $date"
Add-Content -Path $logPath -Value "Software: $nome"
Add-Content -Path $logPath -Value "Version: $versao"

# Envia o arquivo
robocopy "$fullpath" \\$hostname\c$\Temp\ "$instalador" /R:3 /W:10 /J /V /ETA /TEE /LOG+:$logPath
robocopy "$fullpath" \\$hostname\c$\Temp\ "GIAST.sdb" /R:3 /W:10 /J /V /ETA /TEE /LOG+:$logPath

# Inicia a instalação
Add-Content -Path $logPath -Value "Inicia a instalação do $fullname"
$process = psexec \\$hostname -s cmd /c C:\temp\$instalador /silent
Add-Content -Path $logPath -Value "Return: $process"

# Faz a configuração
Add-Content -Path $logPath -Value "Inicia a instalação do banco do $fullname"
$process = psexec \\$hostname -s cmd /c "C:\Program Files (x86)\GIAST3\FBInstall\Firebird-2.1.4.18393_0_Win32.exe" /silent
Add-Content -Path $logPath -Value "Return: $process"
Add-Content -Path $logPath -Value "Inicia a configuração do $fullname"
$process = psexec \\$hostname -s cmd /c sdbinst "C:\temp\GIAST.sdb"
Add-Content -Path $logPath -Value "Return: $process"

# Coloca a permissão na pasta
$process = psexec \\$hostname -s cmd /c cacls "$pasta" /E /P Users:F
Add-Content -Path $logPath -Value "Permission: $process"

# Inicia os serviços
$process = psexec \\$hostname -s cmd /c net start FirebirdGuardianDefaultInstance
Add-Content -Path $logPath -Value "Permission: $process"
$process = psexec \\$hostname -s cmd /c net start FirebirdServerDefaultInstance
Add-Content -Path $logPath -Value "Permission: $process"

# Finaliza o Log
$date = Get-Date
Add-Content -Path $logPath -Value "Complete: $date"

# Informa o analista que usou o script
$myhostname = get-content env:computername
$msg = "A instalação do $nome, no computador $hostname, foi concluida, mais detalhes disponiveis em $logPath"
Invoke-WmiMethod ` -Path Win32_Process ` -Name Create ` -ArgumentList "msg * $msg" ` -ComputerName $myhostname


