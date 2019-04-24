# GUI - Alterar Senha do AD
# Ivo Dias

# Recebe a biblioteca
Add-Type -assembly System.Windows.Forms
# Carrega o modulo
Import-Module ActiveDirectory
# Cria o formulario principal
$GUI = New-Object System.Windows.Forms.Form
# Configura o formulario
$GUI.Text ='TI - Alterar Senha do AD' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario
$GUI.StartPosition = 'CenterScreen' # Inicializa no centro da tela

# Recebe as credenciais
$userADM = $env:UserName
$userADM = get-aduser -identity $useradm
$userfirst = $userADM.givenName
$userlast = $userADM.Surname
$Domain = (Get-ADDomain).DNSRoot
$userAdm = "$Domain\adm.$userfirst$userlast"
$CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm

# Cria a Label com o texto
$lblTexto = New-Object System.Windows.Forms.Label
$lblTexto.Text = "Usuario:"
$lblTexto.Location  = New-Object System.Drawing.Point(0,10)
$lblTexto.AutoSize = $true
$GUI.Controls.Add($lblTexto)

# Cria a caixa de texto
$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Width = 300
$TextBox.Location  = New-Object System.Drawing.Point(60,10)
$GUI.Controls.Add($TextBox)

# Cria o botao
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(400,10)
$Button.Size = New-Object System.Drawing.Size(120,23)
$Button.Text = "Mudar Senha"
$GUI.Controls.Add($Button)

# Cria label de retorno 1 linha
$lblResposta = New-Object System.Windows.Forms.Label
$lblResposta.Text = ""
$lblResposta.Location  = New-Object System.Drawing.Point(0,40)
$lblResposta.AutoSize = $true
$GUI.Controls.Add($lblResposta)

# Cria label de retorno 1 linha
$lbl2linha = New-Object System.Windows.Forms.Label
$lbl2linha.Text = ""
$lbl2linha.Location  = New-Object System.Drawing.Point(0,55)
$lbl2linha.AutoSize = $true
$GUI.Controls.Add($lbl2linha)

# Cria o evento do botao
$Button.Add_Click(
    {
        # Tenta fazer o desbloqueio da conta
        try {
            $Conta = $TextBox.Text
            $Password = Get-Date -Format Sin@mmss
            Set-ADAccountPassword -Identity $Conta -NewPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force) -Credential $CredDomain
            Set-ADUser -Identity $Conta -changepasswordatlogon $true -Credential $CredDomain
            Unlock-ADAccount -Identity $Conta -Credential $CredDomain
            $resposta = "A senha da $conta foi alterada"
            $Linha2 = "Nova senha: $Password"
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao desbloquear a conta|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        $lblResposta.Text =  $resposta
        $lbl2linha.Text = $Linha2
        $TextBox.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()