# Relatorio Office - Email
# Ivo Dias 
# Adaptado de https://www.sconstantinou.com/office-365-powershell-license-report/

function SendMail-OfficeReport 
{
    param 
    (
        [parameter(position = 0, Mandatory = $True)]
        $Conta,
        [parameter(position = 1)]
        $Dominio = "Dominio",
        [parameter(position = 2)]
        $Empresa = "Empresa"
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

        # Configura a empresa
        <#$SKU = $Empresa + ":ENTERPRISEPACK"
        $DuplicateLicenseUsers = (Get-MsolUser -All |
                where {$_.isLicensed -eq "TRUE" -and $_.Licenses.AccountSKUID -eq "$SKU" -and
                $_.Licenses.AccountSKUID -eq "$SKU"}).UserPrincipalName#>

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