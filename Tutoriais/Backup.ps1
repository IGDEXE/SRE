# Ivo Dias
# Script para fazer backup 

# Recebe o caminho da pasta
$pasta = Read-Host "Informe o caminho da pasta que quer fazer o backup[Ex: C:\Fotos] "

# Recebe o local onde quer fazer o backup
$destino = Read-Host "Informe o caminho de onde quer salvar o backup[Ex: C:\Backup] "

# Faz o backup pelo Robocopy
Robocopy $pasta $destino /e /log:"$destino\Log.txt" /R:3 /W:5 /V /ETA

# Mostra mensagem ao completar
Write-Host "O procedimento foi finalizado, mais detalhes em: $destino\Log.txt"