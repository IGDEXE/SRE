# GUI - Remover usuario em um grupo do AD
# Ivo Dias

# Recebe a biblioteca
Add-Type -assembly System.Windows.Forms
# Carrega o modulo
Import-Module ActiveDirectory
# Cria o formulario principal
$GUI = New-Object System.Windows.Forms.Form
# Configura o formulario
$GUI.Text ='TI - Remover usuario em um grupo do AD' # Titulo
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

# Cria a Label com o texto para o grupo
$lblGrupo = New-Object System.Windows.Forms.Label
$lblGrupo.Text = "Grupo:"
$lblGrupo.Location  = New-Object System.Drawing.Point(0,30)
$lblGrupo.AutoSize = $true
$GUI.Controls.Add($lblGrupo)

# Cria a caixa de texto para o grupo
$txtGrupo = New-Object System.Windows.Forms.TextBox
$txtGrupo.Width = 300
$txtGrupo.Location  = New-Object System.Drawing.Point(60,30)
$GUI.Controls.Add($txtGrupo)

# Cria o botao
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(400,20)
$Button.Size = New-Object System.Drawing.Size(120,23)
$Button.Text = "Remover"
$GUI.Controls.Add($Button)

# Cria label de retorno 1 linha
$lblResposta = New-Object System.Windows.Forms.Label
$lblResposta.Text = ""
$lblResposta.Location  = New-Object System.Drawing.Point(0,52)
$lblResposta.AutoSize = $true
$GUI.Controls.Add($lblResposta)

# Cria label de retorno 1 linha
$lbl2linha = New-Object System.Windows.Forms.Label
$lbl2linha.Text = ""
$lbl2linha.Location  = New-Object System.Drawing.Point(0,67)
$lbl2linha.AutoSize = $true
$GUI.Controls.Add($lbl2linha)

# Cria o evento do botao
$Button.Add_Click(
    {
        # Tenta fazer o desbloqueio da conta
        try {
            $Conta = $txtUsuario.Text
            $Grupo = $txtGrupo.Text
            Remove-ADGroupMember -Identity "$Grupo" -Members "$Conta" -Credential $CredDomain
            $resposta = "A $Conta foi removida no grupo $Grupo"
            $Linha2 = ""
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao remover $Conta no grupo $Grupo|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        $lblResposta.Text =  $resposta
        $lbl2linha.Text = $Linha2
        $txtUsuario.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()