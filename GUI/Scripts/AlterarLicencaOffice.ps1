# GUI - Alterar licenca Office
# Ivo Dias

# Criacao do formulario principal
Add-Type -assembly System.Windows.Forms # Recebe a biblioteca
Import-Module MSOnline # Carrega o modulo
$GUI = New-Object System.Windows.Forms.Form # Cria o formulario principal
# Configura o formulario
$GUI.Text ='TI - Alterar licenca Office' # Titulo
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

# Office
$lblOffice = New-Object System.Windows.Forms.Label # Cria a label
$lblOffice.Text = "Licenca:" # Define um texto para ela
$lblOffice.Location  = New-Object System.Drawing.Point(0,30) # Define em qual coordenada da tela vai ser desenhado
$lblOffice.AutoSize = $true # Configura tamanho automatico
$GUI.Controls.Add($lblOffice) # Adiciona ao formulario principal

# Office
$cbxOffice = New-Object System.Windows.Forms.ComboBox
$cbxOffice.Width = 100
$cbxOffice.Location  = New-Object System.Drawing.Point(60,30)
$cbxOffice.Items.Add("Exchange")
$cbxOffice.Items.Add("E1")
$cbxOffice.Items.Add("E1")
$GUI.Controls.Add($cbxOffice)

# Botao para fazer a troca
$Button = New-Object System.Windows.Forms.Button # Cria um botao
$Button.Location = New-Object System.Drawing.Size(190,15) # Define em qual coordenada da tela vai ser desenhado
$Button.Size = New-Object System.Drawing.Size(120,23) # Define o tamanho
$Button.Text = "Modificar" # Define o texto
$GUI.Controls.Add($Button) # Adiciona ao formulario principal

# Label para receber o retorno do procedimento
$lblResposta = New-Object System.Windows.Forms.Label # Cria a label
$lblResposta.Text = "" # Coloca um texto em branco
$lblResposta.Location  = New-Object System.Drawing.Point(0,40) # Define em qual coordenada da tela vai ser desenhado
$lblResposta.AutoSize = $true # Configura tamanho automatico
$GUI.Controls.Add($lblResposta) # Adiciona ao formulario principal

# Label para receber o erro
$lblErro = New-Object System.Windows.Forms.Label # Cria a label
$lblErro.Text = "" # Coloca um texto em branco
$lblErro.Location  = New-Object System.Drawing.Point(0,55) # Define em qual coordenada da tela vai ser desenhado
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
            $Temp = Get-MsolUser -UserPrincipalName $usuario
            $Antigo = $Temp.Licenses.AccountSKUid
            # Configura a nova
            $Novo = $cbxOffice.selectedItem
            if ($Novo -eq "Exchange") { $License = "ATTPS:EXCHANGESTANDARD" }
            if ($Novo -eq "E1") { $License = "ATTPS:STANDARDPACK" }
            if ($Novo -eq "E3") { $License = "ATTPS:ENTERPRISEPACK" }
            Set-MsolUserLicense -UserPrincipalName "$usuarioMail" -AddLicenses "$License" -RemoveLicenses "$Antigo" # Faz a troca
            $resposta = "O licenciamento da conta $usuario foi alterado de $Antigo para $Novo" # Utiliza a label para notificar na tela o resultado
        }
        # Caso encontre algum erro ao fazer o procedimento, retorna
        catch {
            $ErrorMessage = $_.Exception.Message # Recebe a mensagem de erro
            $resposta = "|Ocorreu um erro ao desbloquear a conta|" # Utiliza a label para notificar na tela o resultado
            $lblErro.Text = "|Erro: $ErrorMessage|" # Utiliza a segunda label para mostrar exatamente qual o erro
        }
        $lblResposta.Text =  $resposta # Exibe na tela o retorno do procedimento
        $txtUsuario.Text = "" # Limpa a caixa de texto, para um proximo uso
    }
)

# Inicia o formulario
$GUI.ShowDialog() # Desenha na tela todos os componentes adicionados ao formulario