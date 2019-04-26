# GUI - Adicionar usuario em uma sala do Office
# Ivo Dias

# Criacao do formulario principal
Add-Type -assembly System.Windows.Forms # Recebe a biblioteca
Import-Module MSOnline # Carrega o modulo
$GUI = New-Object System.Windows.Forms.Form # Cria o formulario principal
# Configura o formulario
$GUI.Text ='TI - Adicionar usuario em uma sala do Office' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario
$GUI.StartPosition = 'CenterScreen' # Inicializa no centro da tela

# Recebe a credencial
$userADM = $env:UserName # Recebe o usuario
$userADM += '@sinqia.com.br' # Configura o e-mail
$LiveCred = Get-Credential -Message "Informe as credenciais de Administrador do Office 365" -UserName $userADM

# Conecta no Office
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.outlook.com/powershell/' -Credential $LiveCred -Authentication Basic -AllowRedirection 
    Import-PSSession $Session # Importa a secao
    Connect-MsolService -Credential $LiveCred # Conecta a secao
}
catch {
    # Cria uma mensagem de erro
    Add-Type -AssemblyName PresentationFramework
    $ErrorMessage = $_.Exception.Message # Recebe a mensagem de erro
    $Erro = "$ErrorMessage" # Utiliza a label para notificar na tela o resultado
    [System.Windows.MessageBox]::Show("$Erro",'Erro ao acessar','OK','Error')
    exit
}

# Usuario
$lblUsuario = New-Object System.Windows.Forms.Label # Cria a label
$lblUsuario.Text = "Usuario:" # Define um texto para ela
$lblUsuario.Location  = New-Object System.Drawing.Point(0,10) # Define em qual coordenada da tela vai ser desenhado
$lblUsuario.AutoSize = $true # Configura tamanho automatico
$GUI.Controls.Add($lblUsuario) # Adiciona ao formulario principal

# Caixa de texto para receber Usuario
$txtUsuario = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto
$txtUsuario.Width = 100 # Configura o tamanho
$txtUsuario.Location  = New-Object System.Drawing.Point(60,10) # Define em qual coordenada da tela vai ser desenhado
$GUI.Controls.Add($txtUsuario) # Adiciona ao formulario principal

# Sala
$lblSala = New-Object System.Windows.Forms.Label # Cria a label
$lblSala.Text = "Sala:" # Define um texto para ela
$lblSala.Location  = New-Object System.Drawing.Point(0,30) # Define em qual coordenada da tela vai ser desenhado
$lblSala.AutoSize = $true # Configura tamanho automatico
$GUI.Controls.Add($lblSala) # Adiciona ao formulario principal

# Sala
$cbxSala = New-Object System.Windows.Forms.ComboBox
$cbxSala.Width = 100
$cbxSala.Location  = New-Object System.Drawing.Point(60,30)
$cbxSala.Items.Add("Avengers")
$cbxSala.Items.Add("Jumanji")
$cbxSala.Items.Add("Matrix")
$cbxSala.Items.Add("Star Wars")
$cbxSala.Items.Add("Sonic")
$cbxSala.Items.Add("Super Mario")
$cbxSala.Items.Add("Pacman")
$cbxSala.Items.Add("Tetris")
$cbxSala.Items.Add("SAO 71")
$cbxSala.Items.Add("SAO 72")
$cbxSala.Items.Add("SAO 73")
$cbxSala.Items.Add("Atari")
$cbxSala.Items.Add("Nintendo")
$GUI.Controls.Add($cbxSala)

# Botao para fazer a troca
$Button = New-Object System.Windows.Forms.Button # Cria um botao
$Button.Location = New-Object System.Drawing.Size(190,15) # Define em qual coordenada da tela vai ser desenhado
$Button.Size = New-Object System.Drawing.Size(120,23) # Define o tamanho
$Button.Text = "Adicionar" # Define o texto
$GUI.Controls.Add($Button) # Adiciona ao formulario principal

# Label para receber o retorno do procedimento
$lblResposta = New-Object System.Windows.Forms.Label # Cria a label
$lblResposta.Text = "" # Coloca um texto em branco
$lblResposta.Location  = New-Object System.Drawing.Point(0,55) # Define em qual coordenada da tela vai ser desenhado
$lblResposta.AutoSize = $true # Configura tamanho automatico
$GUI.Controls.Add($lblResposta) # Adiciona ao formulario principal

# Label para receber o erro
$lblErro = New-Object System.Windows.Forms.Label # Cria a label
$lblErro.Text = "" # Coloca um texto em branco
$lblErro.Location  = New-Object System.Drawing.Point(0,70) # Define em qual coordenada da tela vai ser desenhado
$lblErro.AutoSize = $true # Configura tamanho automatico
$GUI.Controls.Add($lblErro) # Adiciona ao formulario principal

# Configura o retorno do botao, quando for utilizado
$Button.Add_Click(
    {
        # Tenta fazer o desbloqueio da conta
        try {
            # Recebe a licenca atual
            $usuario = $txtUsuario.Text
            $usuarioMail = $usuario + '@sinqia.com.br'
            # Verifica se o usuario existe
            if (Get-MsolUser -UserPrincipalName $usuarioMail) {
                # Configura a nova
                $Novo = $cbxSala.selectedItem
                if ($Novo -eq "Avengers") { $sala = "Avengers" }
                if ($Novo -eq "Jumanji") { $sala = "Jumanji" }
                if ($Novo -eq "Matrix") { $sala = "Matrix" }
                if ($Novo -eq "Star Wars") { $sala = "StarWars" }
                if ($Novo -eq "Sonic") { $sala = "Sonic" }
                if ($Novo -eq "Super Mario") { $sala = "SuperMario" }
                if ($Novo -eq "Pacman") { $sala = "Pacman" }
                if ($Novo -eq "Tetris") { $sala = "Tetris" }
                if ($Novo -eq "SAO 71") { $sala = "room.sao71" }
                if ($Novo -eq "SAO 72") { $sala = "room.sao72" }
                if ($Novo -eq "SAO 73") { $sala = "room.sao73" }
                if ($Novo -eq "Atari") { $sala = "Atari" }
                if ($Novo -eq "Nintendo") { $sala = "Nintendo" }
                $sala += '@sinqia.com.br'
                $contatos = (Get-CalendarProcessing -Identity "$sala").BookInPolicy # Recebe os atuais
                $contatos += $usuarioMail # Adiciona o usuario
                Set-CalendarProcessing -Identity $sala -AutomateProcessing AutoAccept -BookInPolicy $contatos # Faz a troca
                $resposta = "O usuario $usuario foi adionado na sala $Novo" # Utiliza a label para notificar na tela o resultado
            } else {
                $resposta = "|Ocorreu um erro ao desbloquear a conta|"
                $lblErro.Text = "|O usuario $usuario nao foi localizado|"
            }
        }
        # Caso encontre algum erro ao fazer o procedimento, retorna
        catch {
            $ErrorMessage = $_.Exception.Message # Recebe a mensagem de erro
            $resposta = "|Ocorreu um erro ao desbloquear a conta|" # Utiliza a label para notificar na tela o resultado
            $lblErro.Text = "|Erro: $ErrorMessage|" # Utiliza a segunda label para mostrar exatamente qual o erro
        }
        $lblResposta.Text = $resposta # Exibe na tela o retorno do procedimento
        $txtUsuario.Text = "" # Limpa a caixa de texto, para um proximo uso
    }
)

# Inicia o formulario
$GUI.ShowDialog() # Desenha na tela todos os componentes adicionados ao formulario