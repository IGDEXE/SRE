# GUI - Novo Usuario Atena
# Ivo Dias

# Recebe a biblioteca
Add-Type -assembly System.Windows.Forms
# Carrega o modulo
Import-Module ActiveDirectory
# Cria o formulario principal
$GUI = New-Object System.Windows.Forms.Form
# Configura o formulario
$GUI.Text ='TI - Novo Usuario Senior' # Titulo
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
$lblSetor = New-Object System.Windows.Forms.Label
$lblSetor.Text = "Setor:"
$lblSetor.Location  = New-Object System.Drawing.Point(0,50)
$lblSetor.AutoSize = $true
$GUI.Controls.Add($lblSetor)

# Cria a caixa de texto dos setores
$cbxSetores = New-Object System.Windows.Forms.ComboBox
$cbxSetores.Width = 300
$cbxSetores.Location  = New-Object System.Drawing.Point(80,50)
$Setores = get-content "\\srsvm030\Scripts\NewUser\ouList.txt"
foreach ($setor in $Setores) {
    # Adiciona como opcao cada um dos setores
    $cbxSetores.Items.Add($setor)
}
$GUI.Controls.Add($cbxSetores)

# Cria a Label com o texto para o usuario de referencia
$lblReferencia = New-Object System.Windows.Forms.Label
$lblReferencia.Text = "Referencia:"
$lblReferencia.Location  = New-Object System.Drawing.Point(0,70)
$lblReferencia.AutoSize = $true
$GUI.Controls.Add($lblReferencia)

# Cria a caixa de texto para o usuario de referencia
$txtReferencia = New-Object System.Windows.Forms.TextBox
$txtReferencia.Width = 300
$txtReferencia.Location  = New-Object System.Drawing.Point(80,70)
$GUI.Controls.Add($txtReferencia)

# Label para as opcoes de grupos
$lblGrupos = New-Object System.Windows.Forms.Label
$lblGrupos.Text = "Grupos:"
$lblGrupos.Location  = New-Object System.Drawing.Point(0,90)
$lblGrupos.AutoSize = $true
$GUI.Controls.Add($lblGrupos)

# Cria a caixa de texto para as opcoes dos grupos
$cbxGrupos = New-Object System.Windows.Forms.CheckedListBox
$cbxGrupos.Width = 300
$cbxGrupos.Location  = New-Object System.Drawing.Point(80,90)
$cbxGrupos.Enabled = $false
$GUI.Controls.Add($cbxGrupos)

# Cria o botao
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(400,27)
$Button.Size = New-Object System.Drawing.Size(120,23)
$Button.Text = "Criar"
$Button.Enabled = $false
$GUI.Controls.Add($Button)

# Cria o botao para verificar
$btnVerificar = New-Object System.Windows.Forms.Button
$btnVerificar.Location = New-Object System.Drawing.Size(400,57)
$btnVerificar.Size = New-Object System.Drawing.Size(120,20)
$btnVerificar.Text = "Verificar"
$GUI.Controls.Add($btnVerificar)

# Cria label de retorno 1 linha
$lblResposta = New-Object System.Windows.Forms.Label
$lblResposta.Text = ""
$lblResposta.Location  = New-Object System.Drawing.Point(0,190)
$lblResposta.AutoSize = $true
$GUI.Controls.Add($lblResposta)

# Cria label de retorno 2 linha
$lbl2linha = New-Object System.Windows.Forms.Label
$lbl2linha.Text = ""
$lbl2linha.Location  = New-Object System.Drawing.Point(0,210)
$lbl2linha.AutoSize = $true
$GUI.Controls.Add($lbl2linha)

# Cria o evento do botao
$btnVerificar.Add_Click(
    {
        # Verifica o usuario de referencia
        try {
            # Configura as referencias de grupo
            $referencia = $txtReferencia.Text
            $cbxGrupos.Enabled = $true
            $userRef = Get-ADUser -Identity $referencia -Properties *
            $Grupos = $userRef.MemberOf
            foreach ($Grupo in $Grupos) {
                # Adiciona como opcao cada um dos setores
                $nomeGrupo = Get-ADGroup -Identity $Grupo
                $nomeGrupo = $nomeGrupo.SamAccountName
                $cbxGrupos.Items.Add($nomeGrupo)
            }
            $Button.Enabled = $true
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um verificar o usuario $referencia|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
    }
)

# Cria o evento do botao
$Button.Add_Click(
    {
        # Faz o procedimento
        try {
            # Recebe os dados
            $nome = $txtNome.Text
            $sobrenome = $txtSobrenome.Text
            $setor = $cbxSetores.selectedItem
            $License = $cbxOffice.selectedItem
            $Grupos = $cbxGrupos.selectedItems
            # Configura os dados
            $userName = "$nome $sobrenome"
            $userAlias = ("$nome.$sobrenome").ToLower()
            $userMail = "$userAlias@sinqia.com.br"
            $OU = "OU=$setor,OU=SeniorSolution,$DomainDC"
            $Password = Get-Date -Format Sin@mmss
            # Cria o usuario
            New-ADUser -SamAccountName $userAlias -Name $userName -GivenName $nome -Surname $sobrenome -EmailAddress $userMail -Path "$OU" -Credential $CredDomain
            Set-ADUser -Identity $userAlias -UserPrincipalName $userMail -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{'msRTCSIP-PrimaryUserAddress'="SIP:$userMail"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{proxyAddresses="SMTP:$userMail","SIP:$userMail"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{targetAddress="SMTP:$userAlias@ATTPS.mail.onmicrosoft.com"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{DisplayName="$userName"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{displayNamePrintable="$userName"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{mailNickname="$userAlias"} -Credential $CredDomain
            Set-ADAccountPassword -Identity $userAlias -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Enabled $true -Credential $CredDomain
            # Configura as referencias de grupo
            foreach ($Item in $cbxGrupos.CheckedItems) {
                $grupo = $Item.ToString()
                Add-ADGroupMember -Identity $grupo -Members $userAlias -Credential $CredDomain
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
        $cbxSetores.Text = ""
        $txtReferencia.Text = ""
        $Button.Enabled = $false
        $cbxGrupos.Enabled = $false
    }
)

# Inicia o formulario
$GUI.ShowDialog()