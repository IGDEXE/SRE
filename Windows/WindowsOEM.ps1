# Ativar Windows 10 OEM
# Ivo Dias

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
function Ativar-Windows10OEM {
    <#
        .SYNOPSIS 
            Faz a ativacao do Windows 10 OEM
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
# Inicia a funcao
Ativar-Windows10OEM