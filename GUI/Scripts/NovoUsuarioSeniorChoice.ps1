# GUI - Novo Usuario Senior
# Ivo Dias

# Recebe a biblioteca
Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
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

<# Selecionar opcoes
$lblInfos = New-Object System.Windows.Forms.Label
$lblInfos.Text = "Informacoes:"
$lblInfos.Location  = New-Object System.Drawing.Point(0,90)
$lblInfos.AutoSize = $true
$GUI.Controls.Add($lblInfos)

# Tipo de contratacao
$cbxContratacao = New-Object System.Windows.Forms.ComboBox
$cbxContratacao.Width = 100
$cbxContratacao.Text = "Contrato"
$cbxContratacao.Location  = New-Object System.Drawing.Point(280,90)
$cbxContratacao.Items.Add("PJ")
$cbxContratacao.Items.Add("CLT")
$cbxContratacao.Items.Add("Estagio")
$GUI.Controls.Add($cbxContratacao)

# Cargo
$cbxCargo = New-Object System.Windows.Forms.ComboBox
$cbxCargo.Width = 100
$cbxCargo.Text = "Cargo"
$cbxCargo.Location  = New-Object System.Drawing.Point(180,90)
$cbxCargo.Items.Add("Alocado")
$cbxCargo.Items.Add("Coordenador")
$cbxCargo.Items.Add("Gerente")
$cbxCargo.Items.Add("Diretor")
$GUI.Controls.Add($cbxCargo)

# Filial
$cbxCargo = New-Object System.Windows.Forms.ComboBox
$cbxCargo.Width = 100
$cbxCargo.Text = "Filial"
$cbxCargo.Location  = New-Object System.Drawing.Point(80,90)
$cbxCargo.Items.Add("Sao Paulo")
$cbxCargo.Items.Add("Rio de Janeiro")
$cbxCargo.Items.Add("Belo Horizonte")
$cbxCargo.Items.Add("Salvador")
$GUI.Controls.Add($cbxCargo) #>

# Cria o evento do botao
$btnVerificar.Add_Click(
    {
        # Verifica o usuario de referencia
        try {
            # Configura as referencias de grupo
            $referencia = $txtReferencia.Text
            $cbxGrupos.Enabled = $true
            # Adiciona os grupos principais
            $gruposPrincipais = get-content "\\srsvm030\Scripts\NewUser\grupoList.txt"
            foreach ($grupo in $gruposPrincipais) {
                # Adiciona como opcao cada um dos setores
                $cbxGrupos.Items.Add($grupo)
            }
            $userRef = Get-ADUser -Identity $referencia -Properties *
            $Grupos = $userRef.MemberOf
            foreach ($Grupo in $Grupos) {
                # Adiciona como opcao cada um dos setores
                $nomeGrupo = Get-ADGroup -Identity $Grupo
                $nomeGrupo = $nomeGrupo.SamAccountName
                $cbxGrupos.Items.Add($nomeGrupo)
            }
                # Configura conforme as opcoes das Combobox
                # Verifica a filial
                <#$grupo = ""
                $escolha = $cbxContratacao.selectedItem
                if ($escolha -eq "Sao Paulo") {
                    $filial = "SP"
                    $cbxGrupos.Items.Add("SINQIASP - Todos Funcionários")
                }
                if ($escolha -eq "Rio de Janeiro") {
                    $filial = "RJ"
                    $cbxGrupos.Items.Add("SINQIARJ - Todos Profissionais")
                }
                if ($escolha -eq "Belo Horizonte") {
                    $filial = "MG"
                    $grupo = ""
                }
                if ($escolha -eq "Salvador") {
                    $filial = "BA"
                    $cbxGrupos.Items.Add("SINQIASAL - Todos Colaboradores")
                }
                # Tipo de contratacao
                $grupo = ""
                $escolha = $cbxContratacao.selectedItem
                if (($escolha -eq "PJ") -and ($filial -eq "SP")) {
                    $cbxGrupos.Items.Add("SINQIASP - Funcionários PJ")
                }
                if (($escolha -eq "CLT") -and ($filial -eq "SP")) {
                    $cbxGrupos.Items.Add("SINQIASP -  Funcionários CLT")
                }
                if (($escolha -eq "Estagio") -and ($filial -eq "SP")) {
                    $cbxGrupos.Items.Add("SINQIASP - Estagiarios")
                }
                # Cargo
                $grupo = ""
                $escolha = $cbxCargo.selectedItem
                if ($escolha -eq "Alocado" -and $filial -eq "SP") {
                    $cbxGrupos.Items.Add("SINQIASP - Funcionários Alocados")
                }
                if (($escolha -eq "Coordenador") -and ($filial -eq "SP")) {
                    $cbxGrupos.Items.Add("SINQIASP - Coordenadores")
                }
                if (($escolha -eq "Coordenador") -and ($filial -eq "RJ")) {
                    $cbxGrupos.Items.Add("SINQIARJ - Coordenadores")
                }
                if (($escolha -eq "Gerente") -and ($filial -eq "SP")) {
                    $cbxGrupos.Items.Add("SINQIASP - Gerentes")
                }
                if (($escolha -eq "Diretor") -and ($filial -eq "SP")) {
                    $cbxGrupos.Items.Add("SINQIASP - Diretores")
                }
                if (($escolha -eq "Diretor") -and ($filial -eq "RJ")) {
                    $cbxGrupos.Items.Add("SINQIARJ - Diretores")
                }#>
            $Button.Enabled = $true
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um verificar o usuario $referencia|"
            $Linha2 = "|Erro: $ErrorMessage|"
            $lblResposta.Text =  $resposta
            $lbl2linha.Text = $Linha2
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
            #$License = $cbxOffice.selectedItem
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
            Set-ADUser -Identity $Conta -changepasswordatlogon $true -Credential $CredDomain
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