# SinqiaAjuda
# Modulo de dialeto Powershell para uso interno na Sinqia
# Ivo Dias

<#
    O objetivo desse projeto eh criar uma linguagem mais simples, porem poderosa, baseada no Powershell
    Os comandos sao escritos em portugues, representam funcoes salvas nesse arquivo

    Comando para versionamento: Get-Date -Format yyyyMMddTHHmmssffff
    Versao atual: 20190408T1402108437

    Detalhes da versao: 20190408T1402108437
    Adicionado o modulo: Ativar-Windows10
    Detalhes: Validacao do licenciamento atual, para evitar o uso desnecessario de chaves KMS

    Detalhes da versao: 20190405T1627366085
    Adicionado o modulo: Ativar-Windows10OEM
    Alterado o nome de Ativar-Windows10 para Ativar-Windows10KMS

    Detalhes da versao: 20190403T1627334068
    Adicionado o modulo: Instalar-GoogleDrive
    Alterado o nome de SendMail-OfficeReport para Email-OfficeLicenciamento

    Detalhes da versao: 20190403T1453502988
    Alterado o padrao de escrita das funcoes
    Alterar-SenhaAD * Adicionado suporte para multiplas contas
    Desbloquear-Conta * Adicionado suporte para multiplas contas
    Adicionar-GrupoAD * Adicionado suporte para multiplos grupos

    Detalhes da versao 20190403T1142518694
    Adicionado o modulo: SendMail-OfficeReport, Deploy-Sinqia
    Configurado os textos de ajuda de cada funcao
    
    Detalhes da versao 20190402T1641064508
    Adicionado o modulo: Desligar-Funcionario

    Detalhes da versao 20190328T1146417362
    Publicacao original, conta com os modulos: Alterar-SenhaAD, Ativar-Office2016, Verificar-Ativacao, Ativar-Windows7, Ativar-Windows10, Copiar-Pasta, Desbloquear-Conta, Adicionar-GrupoAD, Info-VM, Recriar-Usuario, Reparar-PC, Reparar-WindowsUpdate 
#>

# Alterar Senha
function Alterar-SenhaAD {
    <#
        .SYNOPSIS 
            Faz a alteracao da senha de usuario
        .DESCRIPTION
            Alterar-SenhaAD nome.sobrenome
            Alterar-SenhaAD "nome.sobrenome","outro.usuario"
    #>
    param (
        [Parameter(Mandatory=$True)]
        $Contas,
        [parameter(position=1)]
        $senha = "Mudar123"
    )
    try {
        # Recebe as credenciais
        $userADM = $env:UserName
        $userADM = get-aduser -identity $useradm
        $userfirst = $userADM.givenName
        $userlast = $userADM.Surname
        $Domain = (Get-ADDomain).DNSRoot
        $DomainName = (Get-ADDomain).NetBIOSName
        $userAdm = "$Domain\adm.$userfirst$userlast"
        $CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD $DomainName" -UserName $userAdm
        
        # Desbloqueia a conta
        foreach ($Conta in $Contas) {
            Set-ADAccountPassword -Identity $conta -NewPassword (ConvertTo-SecureString -AsPlainText "$senha" -Force) -Credential $CredDomain
            Unlock-ADAccount -Identity $Conta -Credential $CredDomain
            Write-Host "A senha da conta $conta foi alterada para $senha"
        }
    }
    catch {
        Write-Host "Erro ao alterar a senha"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Ativar o Office 2016
function Ativar-Office2016 {
    <#
        .SYNOPSIS 
            Faz a ativacao do Office 2016
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    Param (
        [parameter(position=0)]
        $chave = "XXXXX" # Colocar a sua chave 
    )
    try {
        # Verifica a ativacao do Office
        $OfficeLicense = cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus
        if ($OfficeLicense | Select-String “LICENSE STATUS:  ---LICENSED---” -Quiet) {
            Write-Host "O Office esta ativo"
        }
        else {
            # Faz a ativação do Office
            Write-Host "Atualizando licenciamento"
            cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /inpkey:$chave
            Write-Host "Fazendo a ativacao"
            cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /act
            # Verifica a ativacao do Office
            Clear-Host
            $OfficeLicense = cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus
            if ($OfficeLicense | Select-String “LICENSE STATUS:  ---LICENSED---” -Quiet) { Write-Host "O Office foi ativo com sucesso" }
            else { Write-Host "Abra um Sydle para a ativacao do Office" }
        } 
    }
    catch 
    {
        Write-Host "Erro ao ativar o Office"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Verificar licenciamento
function Verificar-Ativacao {
    <#
        .SYNOPSIS 
            Verifica se o Windows esta licenciado
        .DESCRIPTION
            Verificar-Ativacao SRSSPW187
            Retorno:
            ComputerName Status
            ------------ ------
            srsspw187    Licenciado
    #>
    [CmdletBinding()]
     param(
     [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
     [string]$DNSHostName = $Env:COMPUTERNAME
     )
     process {
        try {
            $wpa = Get-WmiObject SoftwareLicensingProduct -ComputerName $DNSHostName `
            -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" `
            -Property LicenseStatus -ErrorAction Stop
        } 
        catch {
            $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
            $wpa = $null 
        }
        $out = New-Object psobject -Property @{
        ComputerName = $DNSHostName;
        Status = [string]::Empty;
        }
        if ($wpa) {
            :outer foreach($item in $wpa) {
            switch ($item.LicenseStatus) {
            0 {$out.Status = "Nao Licenciado"}
            1 {$out.Status = "Licenciado"; break outer}
            2 {$out.Status = "Fora do periodo de carencia"; break outer}
            3 {$out.Status = "Fora do periodo de tolerancia"; break outer}
            4 {$out.Status = "Nao genuino"; break outer}
            5 {$out.Status = "Notificado"; break outer}
            6 {$out.Status = "Extendido"; break outer}
            default {$out.Status = "Unknown value"}
            }
            }
        }  
        else { $out.Status = $status.Message }
        $out
     }
}

# Ativar Windows 7
function Ativar-Windows7 {
    <#
        .SYNOPSIS 
            Faz a ativacao do Windows 7
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    Param (
        [parameter(position=0)]
        $chave = "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4" # Colocar a sua chave 
    )
    Clear-Host
    # Verifica se o Windows já está ativo
    $validacao = Verificar-Ativacao
    if ($validacao.Status -eq "Licenciado") {
        Write-Host "O Windows ja esta ativo"    
    }
    else {       
        # Utiliza os comandos do SLMGR para fazer a ativacao do Windows 7 com a chave de KMS
        try {
            Write-Host "Carregando os arquivos de licenciamento do Windows"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc
            sleep 10
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk
            Write-Host "Limpando os arquivos antigos de licenciamento"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk $chave
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato
            Write-Host "Fazendo a ativacao"
            sleep 10
            Clear-Host
            $validacao = Verificar-Ativacao
            if ($validacao.Status -eq "Licenciado") {
                Write-Host "O Windows esta ativo"    
            }
            else {
                cscript //B "%windir%\system32\slmgr.vbs" /rearm
                Write-Host "Reinicie o computador e verifique a ativacao"
            }
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "Um erro ocorreu ao tentar ativar o Windows"
            Write-Host "Erro: $ErrorMessage"
        }
    }
}

# Ativar Windows 10 KMS
function Ativar-Windows10KMS {
    <#
        .SYNOPSIS 
            Faz a ativacao do Windows 10
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    Param (
        [parameter(position=0)]
        $chave = "W269N-WFGWX-YVC9B-4J6C9-T83GX" # Colocar a sua chave 
    )
    Clear-Host
    # Verifica se o Windows já está ativo
    $validacao = Verificar-Ativacao
    if ($validacao.Status -eq "Licenciado") {
        Write-Host "O Windows ja esta ativo"    
    }
    else {       
        # Utiliza os comandos do SLMGR para fazer a ativacao do Windows 10 com a chave de KMS
        try {
            Write-Host "Carregando os arquivos de licenciamento do Windows"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc
            sleep 10
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk
            Write-Host "Limpando os arquivos antigos de licenciamento"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk $chave
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato
            Write-Host "Fazendo a ativacao"
            sleep 10
            Clear-Host
            $validacao = Verificar-Ativacao
            if ($validacao.Status -eq "Licenciado") {
                Write-Host "O Windows esta ativo"    
            }
            else {
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /rearm
                Write-Host "Reinicie o computador e verifique a ativacao"
            }
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "Um erro ocorreu ao tentar ativar o Windows"
            Write-Host "Erro: $ErrorMessage"
        }
    }
}

# Copiar Pastas
function Copiar-Pasta {
    <#
        .SYNOPSIS 
            Copia os arquivos de uma pasta, para outra
        .DESCRIPTION
            Copiar-Pasta "Pasta de origem" "Pasta de destino"
    #>
    Param (
        [parameter(position=1)]
        $PastaDestino,
        [parameter(position=0)]
        $PastaBase 
    )
    # Faz a copia dos arquivos
    try {
        Write-Host "Copiando arquivos da pasta '$PastaBase' para a pasta '$PastaDestino'"
        Copy-Item "$PastaBase" -Destination "$PastaDestino" -Recurse -Force
        Write-Host "Finalizado"
    }
    catch {
        Write-Host "Erro ao fazer a copia"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Desbloquear conta
function Desbloquear-Conta {
    <#
        .SYNOPSIS 
            Desbloqueia uma conta do AD
        .DESCRIPTION
            Desbloquear-Conta nome.sobrenome
            Desbloquear-Conta "nome.sobrenome","nome2.sobrenome2"
    #>
    param (
        [Parameter(Mandatory=$True)]
        $Contas
    )
    try {
        # Recebe as credenciais
        $userADM = $env:UserName
        $userADM = get-aduser -identity $useradm
        $userfirst = $userADM.givenName
        $userlast = $userADM.Surname
        $Domain = (Get-ADDomain).DNSRoot
        $DomainName = (Get-ADDomain).NetBIOSName
        $userAdm = "$Domain\adm.$userfirst$userlast"
        $CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD $DomainName" -UserName $userAdm
        
        # Desbloqueia a conta
        foreach ($Conta in $Contas) {
            try {
                Unlock-ADAccount -Identity $Conta -Credential $CredDomain
                Write-Host "A conta $conta foi desbloqueada"
            }
            catch {
                Write-Host "Erro ao desbloquear a conta $conta"
            }      
        }
    }
    catch {
        Write-Host "Erro ao desbloquear a conta"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Adicionar usuario a um grupo do AD
function Adicionar-GrupoAD {
    <#
        .SYNOPSIS 
            Adiciona usuarios num grupo do AD
        .DESCRIPTION
            Adicionar-GrupoAD "Nome do Grupo" nome.sobrenome
            Adicionar-GrupoAD "Nome do Grupo","Nome do Grupo" nome.sobrenome
            Adicionar-GrupoAD "Nome do Grupo" "nome.sobrenome","outro.usuario"
            Adicionar-GrupoAD "Nome do Grupo","Nome do Grupo" "nome.sobrenome","outro.usuario"
    #>
    param (
        [parameter(position=0, Mandatory=$True)]
        $Grupos,
        [parameter(position=1, Mandatory=$True)]
        $Contas
    )
    try {
        # Recebe as credenciais
        $userADM = $env:UserName
        $userADM = get-aduser -identity $useradm
        $userfirst = $userADM.givenName
        $userlast = $userADM.Surname
        $Domain = (Get-ADDomain).DNSRoot
        $userAdm = "$Domain\adm.$userfirst$userlast"
        $CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm
        
        # Dentro dos grupos
        foreach ($Grupo in $Grupos) {
            Write-Host "Configurando o grupo $grupo"
            # Adiciona os usuarios
            foreach ($conta in $Contas) {
                Write-Host "Usuario $conta adicionado ao grupo $grupo"
                Add-ADGroupMember -Identity "$Grupo" -Members "$conta" -Credential $CredDomain
            }
        }
        Write-Host "Procedimento concluido"
    }
    catch {
        Write-Host "Erro ao configurar o grupo"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Informacoes de um servidor de VMs
function Info-VM {
    <#
        .SYNOPSIS 
            Verifica informacoes de VMs em um servidor
        .DESCRIPTION
            Info-VM Servidor
            Info-VM "Servidor1","Servidor2"
    #>
    param (
        [parameter(position=0)]
        $VMServers
    )

    Clear-Host
    Write-Host "Buscando as informacoes das VMs em $VMServers"
    $hash = Get-Date -Format yyyyMMddTHHmmssffff
    $CsvPath = "$env:USERPROFILE\Documents"
    # Tenta acessar a VM e verificar
    foreach ($VMServer in $VMServers) {
        $VMs = Get-VM -ComputerName $VMServer
	    foreach ($VM in $VMs) {
            $VMname = $VM.name
    		    try {
				    Enable-VMResourceMetering -VMName $VMname -ComputerName $VMServer
				    Get-VM $VMname -ComputerName $VMServer | Measure-VM | Select-Object -Property VMName,ComputerName,AvgCPU,AvgRAM,TotalDisk,TotalDiskAllocation,AggregatedDiskDataWritten,AggregatedAverageNormalizedIOPS,AggregatedAverageLatency,AggregatedDiskDataRead,AggregatedNormalizedIOCount | Export-CSV "$CsvPath\ReporteVM.$hash.csv" -NoTypeInformation -Append
				    Measure-VM $VMname -ComputerName $VMServer
    		    }
    		    catch {
                    Write-Host "Erro ao acessar a $VMname"
                    $ErrorMessage = $_.Exception.Message
                    Write-Host "Erro: $ErrorMessage"
    		    }
	    }
    }
}

# Recriar usuario local
function Recriar-Usuario {
    <#
        .SYNOPSIS 
            Recria um usuario local
        .DESCRIPTION
            Para funcionar, eh preciso ter reiniciado o computador
            E o usuario nao fazer logon
            Recriar-Usuario nome.usuario
            Recriar-Usuario nome.usuario computador
            Recriar-Usuario nome.usuario "computador1","computador2"
            Recriar-Usuario "nome.usuario","outro.usuario"
            Recriar-Usuario "nome.usuario","outro.usuario" computador
            Recriar-Usuario "nome.usuario","outro.usuario" "computador1","computador2"
    #>
    param (
        [parameter(position=1)]
        $Computadores = $Env:COMPUTERNAME,
        [parameter(position=0, Mandatory=$True)]
        $Usuarios
    )
    try {
        # Dentro da lista de computadores
        foreach ($computador in $Computadores){
            # Dentro da lista de usuarios
            foreach ($usuario in $Usuarios) {
                Write-Host "Efetuando o procedimento no usuario $usuario dentro do computador $computador"
                # Configura o usuario
                $objUser = New-Object System.Security.Principal.NTAccount($usuario)
                $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
                # Faz o backup
                $hash = Get-Date -Format yyyyMMddTHHmmssffff
                Move-Item -Path "\\$computador\C$\Users\$($objUser.value)" -Destination "\\$computador\C$\Users\$($objUser.value).$hash" -Force
                # Verifica se a .OLD foi criada
                $valida = Test-Path "\\$computador\C$\Users\$($objUser.value).$hash"

                if ($valida -eq  "True") {
                    # Executamos o procedimento de remoção
                    Get-WmiObject -ComputerName $computador win32_userprofile| Where-Object {$_.SID -eq $($strSID.Value)} | ForEach {$_.Delete()} 
                    # Mostra mensagem de encerramento na tela
                    Write-Host "O procedimento foi concluido"    
                }
                else {
                    Write-Host "Um erro ocorreu, e o procedimento não foi feito"
                }
            }
        } 
    }
    catch {
        Write-Host "Erro ao recriar usuario"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Procedimento de desempenho
function Reparar-PC {
    <#
        .SYNOPSIS 
            Faz os procedimentos de desempenho no equipamento
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    # Informa o usuario que o procedimento pode demorar
    Clear-Host
    Write-Host "O procedimento de reparo pode demorar varios minutos"
    Write-Host "Quando ele concluir, a tela vai ficar verde"
    Write-Host "Reinicie quando ele concluir"
    Pause

    # Faz o procedimento
    try {
        $host.ui.RawUI.WindowTitle = "Fazendo reparo"
        sfc /scannow
        Dism.exe /online /cleanup-image /restorehealth
        $host.UI.RawUI.BackgroundColor = [System.ConsoleColor]::Green
        Clear-Host
        $host.ui.RawUI.WindowTitle = "Reparo concluido"
        Pause
        Exit
    }
    catch {
        Write-Host "Erro ao reparar o Windows"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"   
    }
}

# Reparar Windows Update
function Reparar-WindowsUpdate {
    <#
        .SYNOPSIS 
            Faz os procedimentos de reparo do Windows Update no equipamento
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    # Informa o usuario que o procedimento pode demorar
    Clear-Host
    Write-Host "O procedimento de reparo pode demorar varios minutos"
    Write-Host "Quando ele concluir, a tela vai ficar verde"
    # Tenta fazer o reparo dos componentes
    try {
        $host.ui.RawUI.WindowTitle = "Reparando Windows Update"
        # Parando os principais servicos
        Write-Host "Parando os servicos principais"
        Stop-Service -Name "bits" -Force
        Stop-Service -Name "wuauserv" -Force
        Stop-Service -Name "appidsvc" -Force
        Stop-Service -Name "cryptsvc" -Force

        # Apagando os arquivos de configuracao
        Write-Host "Redefinindo as configuracoes"
        Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\" -Recurse -Confirm:$false -Force
        esentutl /d c:\windows\SoftwareDistribution\datastore\datastore.edb

        # Recarrega as DLLs do sistema
        Write-Host "Registrando as DLLs"
        regsvr32.exe atl.dll /s
        regsvr32.exe urlmon.dll /s
        regsvr32.exe mshtml.dll /s
        regsvr32.exe shdocvw.dll /s
        regsvr32.exe browseui.dll /s
        regsvr32.exe jscript.dll /s
        regsvr32.exe vbscript.dll /s
        regsvr32.exe scrrun.dll /s
        regsvr32.exe msxml.dll /s
        regsvr32.exe msxml3.dll /s
        regsvr32.exe msxml6.dll /s
        regsvr32.exe actxprxy.dll /s
        regsvr32.exe softpub.dll /s
        regsvr32.exe wintrust.dll /s
        regsvr32.exe dssenh.dll /s
        regsvr32.exe rsaenh.dll /s
        regsvr32.exe gpkcsp.dll /s
        regsvr32.exe sccbase.dll /s
        regsvr32.exe slbcsp.dll /s
        regsvr32.exe cryptdlg.dll /s
        regsvr32.exe oleaut32.dll /s
        regsvr32.exe ole32.dll /s
        regsvr32.exe shell32.dll /s
        regsvr32.exe initpki.dll /s
        regsvr32.exe wuapi.dll /s
        regsvr32.exe wuaueng.dll /s
        regsvr32.exe wuaueng1.dll /s
        regsvr32.exe wucltui.dll /s
        regsvr32.exe wups.dll /s
        regsvr32.exe wups2.dll /s
        regsvr32.exe wuweb.dll /s
        regsvr32.exe qmgr.dll /s
        regsvr32.exe qmgrprxy.dll /s
        regsvr32.exe wucltux.dll /s
        regsvr32.exe muweb.dll /s
        regsvr32.exe wuwebv.dll /s

        # Ativa os servicos novamente
        Write-Host "Iniciando os servicos principais"
        Start-Service -Name "bits"
        Start-Service -Name "wuauserv"
        Start-Service -Name "appidsvc"
        Start-Service -Name "cryptsvc"
        Pause
        $host.UI.RawUI.BackgroundColor = [System.ConsoleColor]::Green
        Clear-Host
        $host.ui.RawUI.WindowTitle = "Reparo concluido"
    }
    catch {
        Write-Host "Erro ao reparar o Windows Update"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Checklist de Desligamento
function Desligar-Funcionario {
    <#
        .SYNOPSIS 
            Faz o checklist de desligamento, os procedimentos do AD
        .DESCRIPTION
            Desligar-Funcionario nome.sobrenome
            Desligar-Funcionario "nome.sobrenome","outro.usuario"
    #>
    param (
        [Parameter(Mandatory = $True)]
        $userDesligados
    )

        # Carregando modulos
        Write-Host "Carregando modulos..."
        Import-Module MSOnline
        Import-Module ActiveDirectory

    foreach ($userDesligado in $userDesligados) {
        try {
            # Recebe as credenciais de ADM
            $userADM = $env:UserName
            $userADM = Get-ADUser -identity $useradm
            $userfirst = $userADM.givenName
            $userlast = $userADM.Surname
            $Domain = (Get-ADDomain).DNSRoot
            $DomainDC = (Get-ADDomain).DistinguishedName
            $OuDesligados = "OU=Usuários Desativados,$DomainDC"
            $userAdm = "$Domain\adm.$userfirst$userlast"
            $CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm

            # Acesso ao Office
            Write-Host 'Configurando acesso'

            # Recebe a credencial
            $DomainUrl = "@$Domain"
            $userADM = $env:UserName
            $userADM += "@$DomainUrl"
            $LiveCred = Get-Credential -Message "Informe as credenciais de Administrador do Office 365" -UserName $userADM

            # Cria uma nova secao para edicao 
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.outlook.com/powershell/' -Credential $LiveCred -Authentication Basic -AllowRedirection 
            Import-PSSession $Session

           # Conecta a secao
            Connect-MsolService -Credential $LiveCred

            # Confirma exclusao
            $TaSerto = $false
            DO {
               Clear-Host
                # Recebe o usuario
                Write-Host "Usuario que vai ser desligado: $userDesligado"
                Write-Host "Digite 1 para confirmar que o usuario esta correto"
                $escolha = Read-Host "Ou 2 para digitar novamente"
                if ($escolha -eq 1) { $TaSerto = $true }
                if ($escolha -eq 2) { $userDesligado = Read-Host "Informe o usuario desligado (ex: joao.silva)" }
            } While ($TaSerto -eq $false)

            # Configura usuario
            Clear-Host
            $ADUser = $userDesligado
            $userDesligado += "@$DomainUrl"

           Write-Host "Desligando o usuario $userDesligado"
            # Remove a licenca do Office
           $Temp = Get-MsolUser -UserPrincipalName $userDesligado
            $License = $Temp.Licenses.AccountSKUid
            Set-MsolUserLicense -UserPrincipalName "$userDesligado" -RemoveLicenses "$License"
            # Removendo grupos do AD
            $userGroups = Get-ADUser -Identity $ADUser -Properties *
            $userGroups = $userGroups.MemberOf
           foreach ($group in $userGroups) {
                Remove-ADGroupMember -Identity $group -Members $ADUser -Confirm:$false -Credential $CredDomain
            }
            # Move de OU
            $OU = Get-ADUser -Identity $ADUser -Properties *
            $OU = $OU.DistinguishedName
            Move-ADObject -Identity "$OU" -TargetPath "$OuDesligados" -Credential $CredDomain
            # Desativa o usuario
           Set-ADUser -Identity $ADUser -Enabled $false -Credential $CredDomain
            # Exibe informacoes
            Clear-Host
            Write-Host "Usuario $ADUser desligado com sucesso"
            Get-ADUser -Identity $ADUser -Properties * | Select-Object Name, Mail, MemberOf, DistinguishedName, Enabled
            Get-MsolUser -UserPrincipalName "$userDesligado" | Select-Object DisplayName, UsageLocation, @{n = "Licenses Type"; e = { $_.Licenses.AccountSKUid } }
        }
        catch {
           $ErrorMessage = $_.Exception.Message
            Write-Host "Erro ao desligar o usuario $userDesligado"
            Write-Host "Erro: $ErrorMessage"
        }
    }
}

# Deploy this module
function Deploy-ThisTool {
    <#
        .SYNOPSIS 
            Faz a configuracao desse modulo para outros computadores
        .DESCRIPTION
            Deploy-ThisTool computador
            Deploy-ThisTool "computador1","computador2"
    #>
    param (
        [parameter(position=0)]
        $Computadores = $Env:COMPUTERNAME
    )
    # Config
    $server = "" # Informe onde esta o script
    foreach ($computador in $Computadores) {
        try {
            # Copia as pastas
            Write-Host "Fazendo a copia para o computador $computador"
            Copiar-Pasta "$server" "\\$computador\C$\Windows\System32\WindowsPowerShell\v1.0\Modules\"
            Copiar-Pasta "$server" "\\$computador\C$\Windows\SysWOW64\WindowsPowerShell\v1.0\Modules\"

            # Importa o modulo
            psexec \\$computador powershell -executionpolicy unrestricted /c "Import-Module -Name Fable"
        }
        catch {
            Write-Host "Erro ao fazer a configuracao no computador $computador"
            $ErrorMessage = $_.Exception.Message
            Write-Host "Erro: $ErrorMessage"
        }
        Write-Host "---"
    }
    
}

# Relatorio Office - Email
function Email-OfficeLicenciamento {
    <#
        .SYNOPSIS 
            Envia um relatorio sobre o licenciamento do Office por e-mail
        .DESCRIPTION
            SendMail-OfficeReport email@dominio.com.br
            SendMail-OfficeReport "email1@dominio.com.br","email2@dominio.com.br"
    #>
    param 
    (
        [parameter(position = 0, Mandatory = $True)]
        $Conta,
        [parameter(position=1)]
        $Dominio = "contoso.com.br",
        [parameter(position=2)]
        $Empresa = "Contoso"
    )

    try 
    {
        # Notifica o usuario
        Write-host 'Configurando acesso'

        # Importa o modulo
        Import-Module MSOnline

        # Recebe a credencial
        $userADM = $env:UserName
        $userADM += "@$Dominio"
        $LiveCred = Get-Credential -Message "Informe as credenciais de Administrador do Office 365" -UserName $userADM

        # Cria uma nova secao para edicao 
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.outlook.com/powershell/' -Credential $LiveCred -Authentication Basic -AllowRedirection 
        Import-PSSession $Session

        # Conecta a secao
        Connect-MsolService -Credential $LiveCred

        Clear-Host
        Write-Host "Atualizando os dados"
        # Configura e-mail
        $EmailCredentials = $LiveCred
        $To = $Conta
        $From = "$userADM"

        $SKU = $Empresa + ":EXCHANGESTANDARD"
        $TotalP1 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedP1 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableP1 = $TotalP1 - $UsedP1

        $SKU = $Empresa + ":EXCHANGEENTERPRISE"
        $TotalP2 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedP2 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableP2 = $TotalP2 - $UsedP2

        $SKU = $Empresa + ":STANDARDPACK"
        $TotalE1 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedE1 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableE1 = $TotalE1 - $UsedE1

        $SKU = $Empresa + ":ENTERPRISEPACK"
        $TotalE3 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedE3 = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableE3 = $TotalE3 - $UsedE3

        $SKU = $Empresa + ":CRMPLAN2"
        $TotalCRMBasic = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedCRMBasic = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableCRMBasic = $TotalCRMBasic - $UsedCRMBasic

        $SKU = $Empresa + ":CRMSTANDARD"
        $TotalCRMPro = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedCRMPro = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableCRMPro = $TotalCRMPro - $UsedCRMPro

        $SKU = $Empresa + ":CRMINSTANCE"
        $TotalCRMInstance = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedCRMInstance = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableCRMInstance = $TotalCRMInstance - $UsedCRMInstance

        $SKU = $Empresa + ":POWER_BI_STANDARD"
        $TotalBIFree = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedBIFree = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableBIFree = $TotalBIFree - $UsedBIFree

        $SKU = $Empresa + ":POWER_BI_PRO"
        $TotalBIPro = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedBIPro = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableBIPro = $TotalBIPro - $UsedBIPro

        $SKU = $Empresa + ":ATP_ENTERPRISE"
        $TotalATP = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedATP = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableATP = $TotalATP - $UsedATP

        $SKU = $Empresa + ":PROJECTESSENTIALS"
        $TotalProjectEssentials = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedProjectEssentials = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableProjectEssentials = $TotalProjectEssentials - $UsedProjectEssentials

        $SKU = $Empresa + ":PROJECTPREMIUM"
        $TotalProjectPremium = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedProjectPremium = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableProjectPremium = $TotalProjectPremium - $UsedProjectPremium

        $SKU = $Empresa + ":POWERAPPS_VIRAL"
        $TotalPowerApps = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedPowerApps = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailablePowerApps = $TotalPowerApps - $UsedPowerApps

        $SKU = $Empresa + ":STREAM"
        $TotalStream = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ActiveUnits
        $UsedStream = (Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}).ConsumedUnits
        $AvailableStream = $TotalStream - $UsedStream

        $Email = @"
<style>
    body { font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; color:#434242;}
    TABLE { font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TR {border-width: 1px;padding: 10px;border-style: solid;border-color: white; }
    TD {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;padding: 10px;border-style: solid;border-color: white; background-color:#C3DDDB;}
    .colorm {background-color:#58A09E; color:white;}
    .colort{background-color:#58A09E; padding:20px; color:white; font-weight:bold;}
    .colorn{background-color:transparent;}
</style>
<body>

    <h3>Relatorio de licenciamento</h3>

    <table>
        <tr>
            <td class="colorn"></td>
            <td class="colort">Total:</td>
            <td class="colort">Utilizadas:</td>
            <td class="colort">Disponiveis:</td>
        </tr>
        <tr>
            <td class="colorm">Exchange Online Plan1:</td>
            <td style="text-align:center">$TotalP1</td>
            <td style="text-align:center">$UsedP1</td>
            <td style="text-align:center">$AvailableP1</td>
        </tr>
        <tr>
            <td class="colorm">Exchange Online Plan2:</td>
            <td style="text-align:center">$TotalP2</td>
            <td style="text-align:center">$UsedP2</td>
            <td style="text-align:center">$AvailableP2</td>
        </tr>
        <tr>
            <td class="colorm">Office365 Enterprise E1:</td>
            <td style="text-align:center">$TotalE1</td>
            <td style="text-align:center">$UsedE1</td>
            <td style="text-align:center">$AvailableE1</td>
        </tr>
        <tr>
            <td class="colorm">Office365 Enterprise E3:</td>
            <td style="text-align:center">$TotalE3</td>
            <td style="text-align:center">$UsedE3</td>
            <td style="text-align:center">$AvailableE3</td>
        </tr>
        <tr>
            <td class="colorm">Microsoft Dynamics CRM Online Basic:</td>
            <td style="text-align:center">$TotalCRMBasic</td>
            <td style="text-align:center">$UsedCRMBasic</td>
            <td style="text-align:center">$AvailableCRMBasic</td>
        </tr>
        <tr>
            <td class="colorm">Microsoft Dynamics CRM Online Professional:</td>
            <td style="text-align:center">$TotalCRMPro</td>
            <td style="text-align:center">$UsedCRMPro</td>
            <td style="text-align:center">$AvailableCRMPro</td>
        </tr>
        <tr>
            <td class="colorm">Microsoft Dynamics CRM Online Instance:</td>
            <td style="text-align:center">$TotalCRMInstance</td>
            <td style="text-align:center">$UsedCRMInstance</td>
            <td style="text-align:center">$AvailableCRMInstance</td>
        </tr>
        <tr>
            <td class="colorm">Power BI (free):</td>
            <td style="text-align:center">$TotalBIFree</td>
            <td style="text-align:center">$UsedBIFree</td>
            <td style="text-align:center">$AvailableBIFree</td>
        </tr>
        <tr>
            <td class="colorm">Power BI Pro:</td>
            <td style="text-align:center">$TotalBIPro</td>
            <td style="text-align:center">$UsedBIPro</td>
            <td style="text-align:center">$AvailableBIPro</td>
        </tr>
        <tr>
            <td class="colorm">Exchange Online Advance Thread Protection:</td>
            <td style="text-align:center">$TotalATP</td>
            <td style="text-align:center">$UsedATP</td>
            <td style="text-align:center">$AvailableATP</td>
        </tr>
        <tr>
            <td class="colorm">Project Online Essentials:</td>
            <td style="text-align:center">$TotalProjectEssentials</td>
            <td style="text-align:center">$UsedProjectEssentials</td>
            <td style="text-align:center">$AvailableProjectEssentials</td>
        </tr>
        <tr>
            <td class="colorm">Project Online Premium:</td>
            <td style="text-align:center">$TotalProjectPremium</td>
            <td style="text-align:center">$UsedProjectPremium</td>
            <td style="text-align:center">$AvailableProjectPremium</td>
        </tr>
        <tr>
            <td class="colorm">Microsoft Power Apps and Flow:</td>
            <td style="text-align:center">$TotalPowerApps</td>
            <td style="text-align:center">$UsedPowerApps</td>
            <td style="text-align:center">$AvailablePowerApps</td>
        </tr>
        <tr>
            <td class="colorm">Microsoft Stream:</td>
            <td style="text-align:center">$TotalStream</td>
            <td style="text-align:center">$UsedStream</td>
            <td style="text-align:center">$AvailableStream</td>
        </tr>
    </table>
</body>
"@

        # Envia o e-mail
        Write-Host "Enviando o e-mail"
        send-mailmessage `
            -To $To `
            -Subject "Relatorio de licenciamento $(Get-Date -format dd/MM/yyyy)" `
            -Body $Email `
            -BodyAsHtml `
            -Priority high `
            -UseSsl `
            -Port 587 `
            -SmtpServer 'smtp.office365.com' `
            -From $From `
            -Credential $EmailCredentials

        Write-Host "Processo finalizado"
    }
    catch 
    {
        Clear-Host
        Write-Host "Erro ao enviar o e-mail"
        $ErrorMessage = $_.Exception.Message
        Write-Host "Erro: $ErrorMessage"
    }
}

# Instalar Google Drive
function Instalar-GoogleDrive {
    $pastaGoogle = "Caminho onde esta o instalador"
    try {
        # Copia os arquivos de instalacao
        Copiar-Pasta "$pastaGoogle" "C:\INFRA\Google"
        # Faz a instalacao
        Start-Process "C:\INFRA\Google\GoogleDriveFSSetup.exe" -ArgumentList "--silent --desktop_shortcut" -Wait
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Um erro ocorreu ao tentar instalar o Google Drive"
        Write-Host "Erro: $ErrorMessage"
    }
}

# Ativar Windows 10 OEM
function Ativar-Windows10OEM {
    <#
        .SYNOPSIS 
            Faz a ativacao do Windows 10
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    Clear-Host
    # Verifica se o Windows já está ativo
    $validacao = Verificar-Ativacao
    if ($validacao.Status -eq "Licenciado") {
        Write-Host "O Windows ja esta ativo"    
    }
    else {       
        # Utiliza os comandos do SLMGR para fazer a ativacao do Windows 10 com a chave de KMS
        try {
            # Pega a chave OEM
            Write-Host "Recuperando chave OEM"
            $DPK = powershell "(Get-WmiObject -query ‘select * from SoftwareLicensingService’).OA3xOriginalProductKey"
            Write-Host "Carregando os arquivos de licenciamento do Windows"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc
            sleep 10
            Write-Host "Limpando os arquivos antigos de licenciamento"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk
            sleep 10
            Write-Host "Fazendo a ativacao com a chave $DPK"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk $DPK
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato
            sleep 10
            Clear-Host
            $validacao = Verificar-Ativacao
            if ($validacao.Status -eq "Licenciado") {
                Write-Host "O Windows esta ativo"    
            }
            else {
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /rearm
                Write-Host "Reinicie o computador e verifique a ativacao"
            }
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "Um erro ocorreu ao tentar ativar o Windows"
            Write-Host "Erro: $ErrorMessage"
        }
    }
}

# Ativar Windows 10 - Valida KMS e OEM
function Ativar-Windows10 {
    <#
        .SYNOPSIS 
            Faz a ativacao do Windows 10
        .DESCRIPTION
            Nao tem parametros adicionais
    #>
    Clear-Host
    # Verifica se o Windows já está ativo
    $validacao = Verificar-Ativacao
    if ($validacao.Status -eq "Licenciado") {
        # Verifica o tipo de licenciamento do Windows
        # Recebe o licenciamento
        $licenciamento = cscript C:\Windows\System32\slmgr.vbs /dli
        if ($licenciamento | Select-String “VOLUME_KMSCLIENT channel” -Quiet) { 
        # Verifica se a chave OEM existe
        $DPK = powershell "(Get-WmiObject -query ‘select * from SoftwareLicensingService’).OA3xOriginalProductKey"
        if ($DPK -eq "") { 
            Write-Host "Nao existe uma chave OEM nesse equipamento"
            Write-Host "O licenciamento KMS foi mantido"
        }
        else {
            try {
                # Usa a chave OEM para ativar
                Write-Host "Carregando os arquivos de licenciamento do Windows"
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc
                sleep 10
                Write-Host "Limpando os arquivos antigos de licenciamento"
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /cpky
                sleep 10
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk $DPK
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato
                Write-Host "Fazendo a ativacao"
                sleep 10
                Clear-Host
                $validacao = Verificar-Ativacao
                if ($validacao.Status -eq "Licenciado") {
                    Write-Host "Windows OEM - Ativado"    
                }
                else {
                    Write-Host "Nao foi possivel ativar com a chave OEM"
                }
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-Host "Um erro ocorreu ao tentar ativar o Windows"
                Write-Host "Erro: $ErrorMessage"
            }
        }
    }
        if ($licenciamento | Select-String “OEM_DM channel” -Quiet) { 
            Write-Host "Licenciamento OEM" 
        }    
    }
    else {
        # Verifica se a chave OEM existe
        $DPK = powershell "(Get-WmiObject -query ‘select * from SoftwareLicensingService’).OA3xOriginalProductKey"
        if ($DPK -eq "") {
            # Usa o chave KMS para ativar 
            try {
                # Usa a chave OEM para ativar
                Write-Host "Carregando os arquivos de licenciamento do Windows"
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc
                sleep 10
                Write-Host "Limpando os arquivos antigos de licenciamento"
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /cpky
                sleep 10
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato
                Write-Host "Fazendo a ativacao"
                sleep 10
                Clear-Host
                $validacao = Verificar-Ativacao
                if ($validacao.Status -eq "Licenciado") {
                    Write-Host "Windows KMS - Ativado"    
                }
                else {
                    cscript //B "$env:WINDIR\system32\slmgr.vbs" /rearm
                    Write-Host "Nao foi possivel ativar com a chave KMS"
                    Write-Host "Reinicie o computador e verifique a ativacao"
                }
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-Host "Um erro ocorreu ao tentar ativar o Windows"
                Write-Host "Erro: $ErrorMessage"
            }
        }
        else {
            try {
                # Usa a chave OEM para ativar
                Write-Host "Carregando os arquivos de licenciamento do Windows"
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc
                sleep 10
                Write-Host "Limpando os arquivos antigos de licenciamento"
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /cpky
                sleep 10
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk $DPK
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato
                Write-Host "Fazendo a ativacao"
                sleep 10
                Clear-Host
                $validacao = Verificar-Ativacao
                if ($validacao.Status -eq "Licenciado") {
                    Write-Host "Windows OEM - Ativado"    
                }
                else {
                    Write-Host "Nao foi possivel ativar com a chave OEM"
                }
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-Host "Um erro ocorreu ao tentar ativar o Windows"
                Write-Host "Erro: $ErrorMessage"
            }
        }
    }
}

# Desativar usuario
function Desativar-Usuario {
    param (
        [Parameter(Mandatory = $True)]
        $userDesligados
    )
    # Carregando modulos
    Write-Host "Carregando modulos..."
    Import-Module ActiveDirectory
    # Recebe as credenciais de ADM
    $userADM = $env:UserName
    $userADM = get-aduser -identity $useradm
    $userfirst = $userADM.givenName
    $userlast = $userADM.Surname
    $Domain = (Get-ADDomain).DNSRoot
    $DomainName = (Get-ADDomain).NetBIOSName
    $DomainDC = (Get-ADDomain).DistinguishedName
    $ouDesligados = "OU=Usuarios Desativados,$DomainDC"
    $userAdm = "$Domain\adm.$userfirst$userlast"
    $CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD $DomainName" -UserName $userAdm
    # Faz o loop
    foreach ($userDesligado in $userDesligados) {
        # Confirma exclusao
        $TaSerto = $false
        DO {
            Clear-Host
            # Recebe o usuario
            Write-Host "Usuario que vai ser desligado: $userDesligado"
            Write-Host "Digite 1 para confirmar que o usuario esta correto"
            $escolha = Read-Host "Ou 2 para digitar novamente"
            if ($escolha -eq 1) { $TaSerto = $true }
            if ($escolha -eq 2) { $userDesligado = Read-Host "Informe o usuario desligado (ex: joao.silva)" }
        } While ($TaSerto -eq $false)
        # Configura usuario
        Clear-Host
        $ADUser = $userDesligado
        try {
            Write-Host "Desligando o usuario $ADUser"
            # Removendo grupos do AD
            $userGroups = Get-ADUser -Identity $ADUser -Properties *
            $userGroups = $userGroups.MemberOf
            foreach ($group in $userGroups) {
                Remove-ADGroupMember -Identity $group -Members $ADUser -Confirm:$false -Credential $CredDomain
            }
            # Move de OU
            $OU = Get-ADUser -Identity $ADUser -Properties *
            $OU = $OU.DistinguishedName
            Move-ADObject -Identity "$OU" -TargetPath "$ouDesligados" -Credential $CredDomain
            # Desativa o usuario
            Set-ADUser -Identity $ADUser -Enabled $false -Credential $CredDomain
            # Exibe informacoes
            Clear-Host
            Write-Host "Usuario $ADUser desligado com sucesso"
            Get-ADUser -Identity $ADUser -Properties * | Select-Object Name,Mail,MemberOf,DistinguishedName,Enabled
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "Erro ao desligar o usuario $userDesligado"
            Write-Host "Erro: $ErrorMessage"
        }
    }
}

# Verificar licenciamento do Office - Terminal
function Verificar-LicenciamentoOffice {
    param (
        [parameter(position=0)]
        $Path = "C:\TI\Office\Relatorios",
        [parameter(position=1)]
        $Dominio = "Contoso.com.br",
        [parameter(position=2)]
        $Empresa = "CONTOSO"
    )
    # Configuracoes gerais
    $hash = Get-Date -Format yyyyMMddTHHmmssffff
    $LogPath = "$Path\365Relatorio-$hash.txt"
    # Verifica se a pasta de LOG existe
    $valida = Test-Path "$Path"
    if ($valida -ne "True") {
        # Cria a pasta
        $off = mkdir $Path
        Write-Host "Criado a pasta de LOGs: $Path"
    }
    try {
        # Notifica o usuario
        Write-host 'Configurando acesso'
        # Importa o modulo
        Import-Module MSOnline
        # Recebe a credencial
        $userADM = $env:UserName
        $userADM += "@$Dominio"
        $LiveCred = Get-Credential -Message "Informe as credenciais de Administrador do Office 365" -UserName $userADM
        # Cria uma nova secao para edicao 
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.outlook.com/powershell/' -Credential $LiveCred -Authentication Basic -AllowRedirection 
        Import-PSSession $Session
        # Conecta a secao
        Connect-MsolService -Credential $LiveCred
        # Recebe os dados
        $SKU = $Empresa + ":EXCHANGESTANDARD"
        $Ex = Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}
        $SKU = $Empresa + ":STANDARDPACK"
        $E1 = Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}
        $SKU = $Empresa + ":ENTERPRISEPACK"
        $E3 = Get-MsolAccountSku | where {$_.AccountSkuId -eq "$SKU"}
        # Filtra e exibe os dados
        Clear-Host
        Write-Host "Relatorio de licenciamento"
        Add-Content -Path "$LogPath" -Value "Relatorio de licenciamento"
        # Exchange
        $total = $Ex.ActiveUnits
        $livre = $Ex.ConsumedUnits
        $log = "Exchange Online - $livre/$total"
        Write-Host "$log"
        Add-Content -Path "$LogPath" -Value $log
        # E1
        $total = $E1.ActiveUnits
        $livre = $E1.ConsumedUnits
        $log = "E1 - $livre/$total"
        Write-Host "$log"
        Add-Content -Path "$LogPath" -Value $log
        # E3
        $total = $E3.ActiveUnits
        $livre = $E3.ConsumedUnits
        $log = "E3 - $livre/$total"
        Write-Host "$log"
        Add-Content -Path "$LogPath" -Value $log
        Write-host "Tambem disponivel em: $LogPath"
    }
    catch {
        Clear-Host
        $ErrorMessage = $_.Exception.Message
        Write-Host "Um erro ocorreu ao tentar gerar o relatorio"
        Write-Host "Erro: $ErrorMessage"
    }  
}