# GUI - Novo Grupo de Distribuicao
# Ivo Dias

# Formulario principal
Add-Type -assembly System.Windows.Forms # Recebe a biblioteca do formulario
Add-Type -AssemblyName PresentationFramework # Recebea biblioteca das mensagens
Import-Module ActiveDirectory # Carrega o modulo
$GUI = New-Object System.Windows.Forms.Form # Cria o formulario principal
$GUI.Text ='TI - Novo Grupo de Distribuicao' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario
$GUI.StartPosition = 'CenterScreen' # Inicializa no centro da tela

# Credenciais
$userADM = $env:UserName # Recebe o usuario atual
$userADM = get-aduser -identity $useradm # Pega as informacoes dele
$userfirst = $userADM.givenName # Recebe o primeiro nome
$userlast = $userADM.Surname # Recebe o sobrenome
$Domain = (Get-ADDomain).DNSRoot # Recebe o AD atual
$DomainDC = (Get-ADDomain).DistinguishedName # Recebe o nome do AD
$userAdm = "$Domain\adm.$userfirst$userlast" # Configura o ADM com o usuario atual
$CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm # Recebe as credenciais

# Nome
$lblNome = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblNome.Text = "Nome:" # Configura o texto
$lblNome.Location  = New-Object System.Drawing.Point(0,10) # Define a localizacao
$lblNome.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblNome) # Adiciona ao formulario principal
$txtNome = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto para o nome
$txtNome.Width = 300 # Configura o tamanho
$txtNome.Location  = New-Object System.Drawing.Point(80,10) # Define a localizacao
$GUI.Controls.Add($txtNome) # Adiciona ao formulario principal

# Email
$lblEmail = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblEmail.Text = "Email:" # Configura o texto
$lblEmail.Location  = New-Object System.Drawing.Point(0,30) # Define a localizacao
$lblEmail.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblEmail) # Adiciona ao formulario principal
$txtEmail = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto
$txtEmail.Width = 300 # Configura o tamanho
$txtEmail.Location  = New-Object System.Drawing.Point(80,30) # Define a localizacao
$GUI.Controls.Add($txtEmail) # Adiciona ao formulario principal

# Usuarios
$lblUsuarios = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblUsuarios.Text = "Usuarios:" # Configura o texto
$lblUsuarios.Location  = New-Object System.Drawing.Point(0,50) # Define a localizacao
$lblUsuarios.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblUsuarios) # Adiciona ao formulario principal
$txtUsuarios = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto
$txtUsuarios.Size = New-Object System.Drawing.Size(300,100) # Configura o tamanho
$txtUsuarios.Multiline = $true # Ativa a opcao de varias linhas
$txtUsuarios.Location  = New-Object System.Drawing.Point(80,50) # Define a localizacao
$GUI.Controls.Add($txtUsuarios) # Adiciona ao formulario principal

# Botao de Criar
$Button = New-Object System.Windows.Forms.Button # Cria o botao
$Button.Location = New-Object System.Drawing.Size(400,27) # Define a localizacao
$Button.Size = New-Object System.Drawing.Size(120,50) # Configura o tamanho
$Button.Text = "Criar" # Configura o texto
$GUI.Controls.Add($Button) # Adiciona ao formulario principal

# Campo de retorno
$lblResposta = New-Object System.Windows.Forms.Label # Cria a label para receber o retorno
$lblResposta.Text = "" # Coloca o texto inicial vazio
$lblResposta.Location  = New-Object System.Drawing.Point(0,170) # Define a localizacao
$lblResposta.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblResposta) # Adiciona ao formulario principal

# Label de retorno 2 linha
$lbl2linha = New-Object System.Windows.Forms.Label # Cria a label para receber o retorno
$lbl2linha.Text = "" # Coloca o texto inicial vazio
$lbl2linha.Location  = New-Object System.Drawing.Point(0,200) # Define a localizacao
$lbl2linha.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lbl2linha) # Adiciona ao formulario principal

# Botao para criar
$Button.Add_Click(
    {
        # Faz o procedimento
        try {
            $nomeGrupo = $txtNome.Text # Recebe o nome do grupo
            $Email = $txtEmail.Text # Recebe o nome do grupo
            $Email += '@sinqia.com.br' # Configura o e-mail
            $usuariosGrupo = $txtUsuarios.Lines # Recebe os usuarios, um por linha
            $OU = "OU=Distribuicao,OU=Grupos,$DomainDC" # Configura a OU
            New-ADGroup -Path "$OU" -Name "$nomeGrupo" -GroupScope DomainLocal -GroupCategory Distribution -Credential $CredDomain # Cria o grupo
            Set-ADGroup -Identity $nomeGrupo -Replace @{mail="$Email"} -Credential $CredDomain # Define o e-mail
            Set-ADGroup -Identity $nomeGrupo -Replace @{proxyAddresses="SMTP:$Email"} -Credential $CredDomain # Configura o SMTP
            # Mostra na tela
            $resposta = "Criado o grupo $nomeGrupo, com os usuarios:"
            foreach ($usuario in $usuariosGrupo) {
                Add-ADGroupMember -Identity $nomeGrupo -Members $usuario -Credential $CredDomain # Adiciona os usuarios 
            }
            $Linha2 = $usuariosGrupo # Mostra os usuarios cadastrados
        }
        catch {
            # Reporta o erro, caso ocorra
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao criar o grupo $nomeGrupo|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        # Escreve o retorno
        $lblResposta.Text =  $resposta
        $lbl2linha.Text = $Linha2
        # Limpa os campos
        $txtNome.Text = ""
        $txtEmail.Text = ""
        $txtUsuarios.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()