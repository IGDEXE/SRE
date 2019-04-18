# GUI - Desbloquear conta
# Ivo Dias

# Recebe a biblioteca
Add-Type -assembly System.Windows.Forms
# Cria o formulario principal
$GUI = New-Object System.Windows.Forms.Form
# Configura o formulario
$GUI.Text ='TI - Desbloquear Conta' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario

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
$Button.Text = "Desbloquear"
$GUI.Controls.Add($Button)

# Cria label de retorno
$lblResposta = New-Object System.Windows.Forms.Label
$lblResposta.Text = ""
$lblResposta.Location  = New-Object System.Drawing.Point(0,40)
$lblResposta.AutoSize = $true
$GUI.Controls.Add($lblResposta)

# Cria label de erro
$lblErro = New-Object System.Windows.Forms.Label
$lblErro.Text = ""
$lblErro.Location  = New-Object System.Drawing.Point(0,55)
$lblErro.AutoSize = $true
$GUI.Controls.Add($lblErro)

# Cria o evento do botao
$Button.Add_Click(
    {
        # Tenta fazer o desbloqueio da conta
        try {
            $Conta = $TextBox.Text
            Unlock-ADAccount -Identity $Conta -Credential $CredDomain
            $resposta = "A conta $conta foi desbloqueada"
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao desbloquear a conta|"
            $lblErro.Text = "|Erro: $ErrorMessage|"
        }
        $lblResposta.Text =  $resposta
        $TextBox.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()