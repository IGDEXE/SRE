# GUI - Desligar Funcionario
# Ivo Dias

# Recebe a biblioteca
Add-Type -assembly System.Windows.Forms
# Cria o formulario principal
$GUI = New-Object System.Windows.Forms.Form
# Configura o formulario
$GUI.Text ='TI - Desligar Funcionario' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario

# Recebe as credenciais
$userADM = $env:UserName
$userADM = get-aduser -identity $useradm
$userfirst = $userADM.givenName
$userlast = $userADM.Surname
$Domain = (Get-ADDomain).DNSRoot
$DomainDC = (Get-ADDomain).DistinguishedName
$userAdm = "$Domain\adm.$userfirst$userlast"
$CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm

# Cria a Label com o texto
$lblUsuario = New-Object System.Windows.Forms.Label
$lblUsuario.Text = "Usuario:"
$lblUsuario.Location  = New-Object System.Drawing.Point(0,10)
$lblUsuario.AutoSize = $true
$GUI.Controls.Add($lblUsuario)

# Cria a caixa de texto para o usuario
$txtUsuario = New-Object System.Windows.Forms.TextBox
$txtUsuario.Width = 300
$txtUsuario.Location  = New-Object System.Drawing.Point(60,10)
$GUI.Controls.Add($txtUsuario)

# Cria o botao
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(400,10)
$Button.Size = New-Object System.Drawing.Size(120,23)
$Button.Text = "Desligar"
$GUI.Controls.Add($Button)

# Cria label de retorno 1 linha
$lblResposta = New-Object System.Windows.Forms.Label
$lblResposta.Text = ""
$lblResposta.Location  = New-Object System.Drawing.Point(0,52)
$lblResposta.AutoSize = $true
$GUI.Controls.Add($lblResposta)

# Cria label de retorno 2 linha
$lbl2linha = New-Object System.Windows.Forms.Label
$lbl2linha.Text = ""
$lbl2linha.Location  = New-Object System.Drawing.Point(0,67)
$lbl2linha.AutoSize = $true
$GUI.Controls.Add($lbl2linha)

# Cria o evento do botao
$Button.Add_Click(
    {
        # Tenta fazer o desligamento do usuario
        try {
            $Conta = $txtUsuario.Text
            # Removendo grupos do AD
            $userGroups = Get-ADUser -Identity $Conta -Properties *
            $userGroups = $userGroups.MemberOf
            foreach ($group in $userGroups) {
                Remove-ADGroupMember -Identity $group -Members $Conta -Confirm:$false -Credential $CredDomain
            }
            # Move de OU
            $OU = Get-ADUser -Identity $Conta -Properties *
            $OU = $OU.DistinguishedName
            $OuDesligados = "OU=Usuários Desativados,$DomainDC"
            Move-ADObject -Identity "$OU" -TargetPath "$OuDesligados" -Credential $CredDomain
            # Desativa o usuario
            Set-ADUser -Identity $Conta -Enabled $false -Credential $CredDomain
            $resposta = "A conta $Conta foi desligada"
            $Linha2 = ""
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao desligar o usuario $Conta|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        $lblResposta.Text =  $resposta
        $lbl2linha.Text = $Linha2
        $txtUsuario.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()