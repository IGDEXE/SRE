# Ativar Windows 10 OEM
# Ivo Dias
# Verify licensing
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
function Activate-Windows10OEM{
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
        Write-Host "Windows is already licensing"    
    }
    else {       
        # Utiliza os comandos do SLMGR para fazer a ativacao do Windows 10 com a chave de KMS
        try {
            # Pega a chave OEM
            Write-Host "Recovering OEM Key"
            $DPK = powershell "(Get-WmiObject -query ‘select * from SoftwareLicensingService’).OA3xOriginalProductKey"
            Write-Host "Loading the Windows Licensing Files"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /rilc
            sleep 10
            Write-Host "Clearing old licensing files"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /upk
            sleep 10
            Write-Host "Activating with the key: $DPK"
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ipk $DPK
            cscript //B "$env:WINDIR\system32\slmgr.vbs" /ato
            sleep 10
            Clear-Host
            $validacao = Verificar-Ativacao
            if ($validacao.Status -eq "Licenciado") {
                Write-Host "Windows is licensing"    
            }
            else {
                cscript //B "$env:WINDIR\system32\slmgr.vbs" /rearm
                Write-Host "Restart the computer and check the licensing"
            }
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "An error occurred while trying to activate Windows"
            Write-Host "Error: $ErrorMessage"
        }
    }
}

# Inicia o script
Activate-Windows10OEM