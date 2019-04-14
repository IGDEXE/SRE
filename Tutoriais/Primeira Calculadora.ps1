# Autor  : Ivo Dias
# Social : igd753
# E-mail : igd753@outlook.com.br

# Limpa a tela
cls

# Coloca uma mensagem inicial
Write-Host "Nossa primeira calculadora"

# Recebe um número, guarda em uma variável, que vamos chamar de “a”:
[double]$a = Read-Host "Informe um número"

# Recebe um outro número, guarda em uma outra variável, que vamos chamar de “b":
[double]$b = Read-Host "Informe um outro número"

# Efetuamos as operações e mostramos na tela: 
# Soma
[double]$c = $a + $b # Faz a conta
Write-Host "O resultado da soma é $c" # Mostra na tela

# Subtração
[double]$c = $a - $b # Faz a conta
Write-Host "O resultado da subtração é $c" # Mostra na tela

# Multiplicação
[double]$c = $a * $b # Faz a conta
Write-Host "O resultado da multiplicação é $c" # Mostra na tela

# Divisão
[double]$c = $a / $b # Faz a conta
Write-Host "O resultado da divisão é $c" # Mostra na tela