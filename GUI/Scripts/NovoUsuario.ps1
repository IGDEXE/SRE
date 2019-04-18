# GUI - Novo Usuario
# Ivo Dias

# Recebe a biblioteca
Add-Type -assembly System.Windows.Forms
# Carrega o modulo
Import-Module ActiveDirectory
# Cria o formulario principal
$GUI = New-Object System.Windows.Forms.Form
# Configura o formulario
$GUI.Text ='TI - Novo Usuario' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario
$GUI.StartPosition = 'CenterScreen' # Inicializa no centro da tela

# Recebe as credenciais
$userADM = $env:UserName
$userADM = get-aduser -identity $useradm
$userfirst = $userADM.givenName
$userlast = $userADM.Surname
$Domain = (Get-ADDomain).DNSRoot
$DomainDC = (Get-ADDomain).DistinguishedName
$userAdm = "$Domain\adm.$userfirst$userlast"
$CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm

# Cria a Label com o texto para o nome
$lblNome = New-Object System.Windows.Forms.Label
$lblNome.Text = "Nome:"
$lblNome.Location  = New-Object System.Drawing.Point(0,10)
$lblNome.AutoSize = $true
$GUI.Controls.Add($lblNome)

# Cria a caixa de texto para o nome
$txtNome = New-Object System.Windows.Forms.TextBox
$txtNome.Width = 300
$txtNome.Location  = New-Object System.Drawing.Point(80,10)
$GUI.Controls.Add($txtNome)

# Cria a Label com o texto para o sobrenome
$lblSobrenome = New-Object System.Windows.Forms.Label
$lblSobrenome.Text = "Sobrenome:"
$lblSobrenome.Location  = New-Object System.Drawing.Point(0,30)
$lblSobrenome.AutoSize = $true
$GUI.Controls.Add($lblSobrenome)

# Cria a caixa de texto para o sobrenome
$txtSobrenome = New-Object System.Windows.Forms.TextBox
$txtSobrenome.Width = 300
$txtSobrenome.Location  = New-Object System.Drawing.Point(80,30)
$GUI.Controls.Add($txtSobrenome)

# Cria a Label com o texto para o usuario de referencia
$lblReferencia = New-Object System.Windows.Forms.Label
$lblReferencia.Text = "Referencia:"
$lblReferencia.Location  = New-Object System.Drawing.Point(0,50)
$lblReferencia.AutoSize = $true
$GUI.Controls.Add($lblReferencia)

# Cria a caixa de texto para o usuario
$txtReferencia = New-Object System.Windows.Forms.TextBox
$txtReferencia.Width = 300
$txtReferencia.Location  = New-Object System.Drawing.Point(80,50)
$GUI.Controls.Add($txtReferencia)

# Cria o botao
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(400,27)
$Button.Size = New-Object System.Drawing.Size(120,23)
$Button.Text = "Criar"
$GUI.Controls.Add($Button)

# Cria label de retorno 1 linha
$lblResposta = New-Object System.Windows.Forms.Label
$lblResposta.Text = ""
$lblResposta.Location  = New-Object System.Drawing.Point(0,70)
$lblResposta.AutoSize = $true
$GUI.Controls.Add($lblResposta)

# Cria label de retorno 2 linha
$lbl2linha = New-Object System.Windows.Forms.Label
$lbl2linha.Text = ""
$lbl2linha.Location  = New-Object System.Drawing.Point(0,87)
$lbl2linha.AutoSize = $true
$GUI.Controls.Add($lbl2linha)

# Cria o evento do botao
$Button.Add_Click(
    {
        # Faz o procedimento
        try {
            # Recebe os dados
            $nome = $txtNome.Text
            $sobrenome = $txtSobrenome.Text
            $referencia = $txtReferencia.Text
            # Configura os dados
            $userName = "$nome $sobrenome"
            $userAlias = ("$nome.$sobrenome").ToLower()
            $userMail = "$userAlias@sinqia.com.br"
            $OU = "OU=Users,$DomainDC"
            $Password = Get-Date -Format Pass@mmss
            # Cria o usuario
            New-ADUser -SamAccountName $userAlias -Name $userName -GivenName $nome -Surname $sobrenome -EmailAddress $userMail -Path "$OU" -Credential $CredDomain
            Set-ADUser -Identity $userAlias -UserPrincipalName $userMail -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{'msRTCSIP-PrimaryUserAddress'="SIP:$userMail"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{proxyAddresses="SMTP:$userMail","SIP:$userMail"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{targetAddress="SMTP:$userMail"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{DisplayName="$userName"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{displayNamePrintable="$userName"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{mailNickname="$userAlias"} -Credential $CredDomain
            Set-ADAccountPassword -Identity $userAlias -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Enabled $true -Credential $CredDomain
            # Configura as referencias de grupo
            $userRef = Get-ADUser -Identity $referencia -Properties *
            $memberOf = $userRef.MemberOf
            foreach ($group in $memberof) {
                Add-ADGroupMember -Identity $group -Members $userAlias -Credential $CredDomain
            }
            # Mostra na tela
            $resposta = "Usuario $userName criado com sucesso"
            $Linha2 = "Senha: $Password"
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao criar o usuario $userName|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        $lblResposta.Text =  $resposta
        $lbl2linha.Text = $Linha2
        $txtNome.Text = ""
        $txtSobrenome.Text = ""
        $txtReferencia.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()