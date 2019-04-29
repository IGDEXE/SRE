# GUI - Alterar Senha do AD
# Ivo Dias

# Formulario principal
Add-Type -assembly System.Windows.Forms # Recebe a biblioteca
Import-Module ActiveDirectory # Carrega o modulo
$GUI = New-Object System.Windows.Forms.Form # Cria o formulario principal
$GUI.Text ='TI - Alterar Senha do AD' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario
$GUI.StartPosition = 'CenterScreen' # Inicializa no centro da tela

# Recebe as credenciais
$userADM = $env:UserName # Recebe o usuario atual
$userADM = get-aduser -identity $useradm # Pega as informacoes dele no AD
$userfirst = $userADM.givenName # Recebe o primeiro nome
$userlast = $userADM.Surname # Recebe o sobrenome
$Domain = (Get-ADDomain).DNSRoot # Verifica o AD
$userAdm = "$Domain\adm.$userfirst$userlast" # Configura o login da credencial
$CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm # Armazena as credenciais

# Usuario
$lblTexto = New-Object System.Windows.Forms.Label # Cria uma label para o usuario
$lblTexto.Text = "Usuario:" # Define o texto padrao
$lblTexto.Location  = New-Object System.Drawing.Point(0,10) # Define a localizacao
$lblTexto.AutoSize = $true # Tamanho automatico
$GUI.Controls.Add($lblTexto) # Adiciona ao formulario principal
$TextBox = New-Object System.Windows.Forms.TextBox # Cria uma caixa de texto para o usuario
$TextBox.Width = 300 # Define um tamanho
$TextBox.Location  = New-Object System.Drawing.Point(60,10) # Define a localizacao
$GUI.Controls.Add($TextBox) # Adiciona ao formulario principal

# Botao
$Button = New-Object System.Windows.Forms.Button # Cria um botao
$Button.Location = New-Object System.Drawing.Size(400,10) # Define a localizacao
$Button.Size = New-Object System.Drawing.Size(120,23) # Define o tamanho
$Button.Text = "Mudar Senha" # Define o texto
$GUI.Controls.Add($Button) # Adiciona ao formulario principal

# Cria label de retorno 1 linha
$lblResposta = New-Object System.Windows.Forms.Label # Cria uma label para o retorno padrao
$lblResposta.Text = "" # Define o texto inicial como vazio
$lblResposta.Location  = New-Object System.Drawing.Point(0,40) # Define a localizacao
$lblResposta.AutoSize = $true # Define o tamanho automatico
$GUI.Controls.Add($lblResposta) # Adiciona ao formulario

# Cria label de retorno 1 linha
$lbl2linha = New-Object System.Windows.Forms.Label # Cria uma label para erros
$lbl2linha.Text = "" # Define o texto inicial como vazio
$lbl2linha.Location  = New-Object System.Drawing.Point(0,55) # Define a localizacao
$lbl2linha.AutoSize = $true # Define o tamanho automatico
$GUI.Controls.Add($lbl2linha) # Adiciona ao formulario

# Cria o evento do botao
$Button.Add_Click(
    {
        # Tenta fazer o desbloqueio da conta
        try {
            $Conta = $TextBox.Text # Recebe a conta
            $Password = Get-Date -Format Sin@mmss # Gera uma nova senha
            Set-ADAccountPassword -Identity $Conta -NewPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force) -Credential $CredDomain # Faz a troca da senha
            Set-ADUser -Identity $Conta -changepasswordatlogon $true -Credential $CredDomain # Configura para o usuario ter que trocar a senha no proximo login
            Unlock-ADAccount -Identity $Conta -Credential $CredDomain # Desbloqueia a conta
            # Carrega o retorno
            $resposta = "A senha da $conta foi alterada" 
            $Linha2 = "Nova senha: $Password"
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao desbloquear a conta|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        # Exibe na tela
        $lblResposta.Text =  $resposta
        $lbl2linha.Text = $Linha2
        $TextBox.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()