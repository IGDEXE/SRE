# GUI - Adicionar usuario em um grupo do AD
# Ivo Dias

# Formulario principal
Add-Type -assembly System.Windows.Forms # Recebe a biblioteca
Import-Module ActiveDirectory # Carrega o modulo
$GUI = New-Object System.Windows.Forms.Form # Cria o formulario principal
$GUI.Text ='TI - Adicionar usuario em um grupo do AD' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario
$GUI.StartPosition = 'CenterScreen' # Inicializa no centro da tela

# Credencial
$userADM = $env:UserName # Recebe o usuario atual
$userADM = get-aduser -identity $useradm # Pega as informacoes dele no AD
$userfirst = $userADM.givenName # Recebe o primeiro nome
$userlast = $userADM.Surname # Recebe o sobrenome
$Domain = (Get-ADDomain).DNSRoot # Verifica o AD
$userAdm = "$Domain\adm.$userfirst$userlast" # Configura o login da credencial
$CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm # Armazena as credenciais

# Usuario
$lblUsuario = New-Object System.Windows.Forms.Label # Cria uma label para o usuario
$lblUsuario.Text = "Usuario:" # Define o texto
$lblUsuario.Location  = New-Object System.Drawing.Point(0,10) # Define a localizacao
$lblUsuario.AutoSize = $true # Configura o aumento de tamanho
$GUI.Controls.Add($lblUsuario) # Adiciona ao formulario
$txtUsuario = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto para o usuario
$txtUsuario.Width = 300 # Define um tamanho
$txtUsuario.Location  = New-Object System.Drawing.Point(60,10) # Define a localizacao
$GUI.Controls.Add($txtUsuario) # Adiciona ao formulario

# Grupo
$lblGrupo = New-Object System.Windows.Forms.Label # Cria a label para o grupo
$lblGrupo.Text = "Grupo:" # Define o texto
$lblGrupo.Location  = New-Object System.Drawing.Point(0,30) # Define a localizacao
$lblGrupo.AutoSize = $true # Configura o aumento de tamanho
$GUI.Controls.Add($lblGrupo) # Adiciona ao formulario
$txtGrupo = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto para o grupo
$txtGrupo.Width = 300 # Define um tamanho
$txtGrupo.Location  = New-Object System.Drawing.Point(60,30) # Define a localizacao
$GUI.Controls.Add($txtGrupo) # Adiciona ao formulario

# Botao
$Button = New-Object System.Windows.Forms.Button # Cria um botao
$Button.Location = New-Object System.Drawing.Size(400,20) # Define a localizacao 
$Button.Size = New-Object System.Drawing.Size(120,23) # Define o tamanho
$Button.Text = "Adicionar" # Define o texto
$GUI.Controls.Add($Button) # Adiciona ao formulario

# Retorno padrao
$lblResposta = New-Object System.Windows.Forms.Label # Cria uma label para o retorno padrao
$lblResposta.Text = "" # Define o texto inicial como vazio
$lblResposta.Location  = New-Object System.Drawing.Point(0,52) # Define a localizacao
$lblResposta.AutoSize = $true # Configura tamanho automatico
$GUI.Controls.Add($lblResposta) # Adiciona ao formulario

# Erro
$lbl2linha = New-Object System.Windows.Forms.Label # Cria uma label para erros
$lbl2linha.Text = "" # Define o texto inicial como vazio
$lbl2linha.Location  = New-Object System.Drawing.Point(0,67) # Define a localizacao
$lbl2linha.AutoSize = $true # Define o tamanho automatico
$GUI.Controls.Add($lbl2linha) # Adiciona ao formulario

# Evento do botao
$Button.Add_Click(
    {
        # Tenta fazer o desbloqueio da conta
        try {
            $Conta = $txtUsuario.Text # Recebe a conta 
            $Grupo = $txtGrupo.Text # Recebe o grupo
            Add-ADGroupMember -Identity "$Grupo" -Members "$Conta" -Credential $CredDomain # Faz o desbloqueio no AD
            $resposta = "A $Conta foi adicionada no grupo $Grupo" # Retorna o resultado
            $Linha2 = ""
        }
        # Se der erro
        catch {
            $ErrorMessage = $_.Exception.Message # Recebe a mensagem de erro
            $resposta = "|Ocorreu um erro ao adicionar $Conta no grupo $Grupo|" # Carrega a mensagem de erro
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        # Escreve na tela o retorno
        $lblResposta.Text = $resposta
        $lbl2linha.Text = $Linha2
        $txtUsuario.Text = "" # Limpa a caixa de texto
        $txtGrupo.Text = "" # Limpa a caixa de texto
    }
)

# Inicia o formulario
$GUI.ShowDialog()