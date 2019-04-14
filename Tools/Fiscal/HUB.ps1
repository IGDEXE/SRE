# Instalador de Fiscais remoto
# AUTHOR  : Ivo Dias 
# VERSION : 22.10.18.10.MSC

 # Functions library:
 # Sets up the path
 function inicio
 {
    cls # Cleans the screen
    # Writes the menu options
    Write-Host "... Tax Program Installer"
    $Installpath = "C:\Fiscal" # Sets up the mapped shared folder
    Remove-Item -Path "$Installpath\Config\user.SID"
    Add-Content -Path "$Installpath\Config\user.SID" -Value $Installpath
    psexec_check # Starts the program
 }
 
# Verifies Psexec
function psexec_check
{
    $valida = "C:\Windows\System32\PsExec.exe" # Checks if psexec exists
    if ($valida -eq  "True")
    {
        INI # Goes to the menu    
    }
    else 
    {
        # Copies psexec to the folder
        Copy-Item -Path "$Installpath\Config\psexec.exe" -Destination "C:\Windows\System32\" -Force
        INI # Goes to the menu  
    }
}

# Starting screen
function INI
{
    cls # Cleans the screen
    # Coloca o titulo no Script
    $host.ui.RawUI.WindowTitle = "Tax Program Installer"
    Write-Host "... Tax Program Installer"
    Write-Host ".. Remember to register any procedures in the ticket"
    Test_Connection # Tests the connection
}

# Tests the connection
function Test_Connection
{
    try 
    {
        # Receives the hostname where the action will be performed
        Write-Host ".. Before starting a procedure, checking if the host is reachable"
        $hostname = Read-Host ".. Inform the hostname or IP of the workstation"

        # Tests connection
        if (Test-Connection -ComputerName $hostname -Quiet)
        {
            Remove-Item -Path "$Installpath\Config\Hostname.SID" # Deletes the file where the hostname is stored
            Add-Content -Path "$Installpath\Config\Hostname.SID" -Value $hostname # Creates the file where the hostname is stored
            menu_principal # Opens the menu
        }
        else
        {
            # Shows the error message
            Write-Host ".. There's been an error accessing $hostname"
            Write-Host ".. Verify if it's typed correctly or connected to a network"
            pause
            INI # Back to the menu
        }
    }
    catch
    {
            Write-Host ".. An error occurred while receiving the hostname"
            pause
            INI # Back to the menu
    }
}

function menu_principal
{
    # Limpa a tela
    cls
    # Mostra as opções
    Write-Host "... Tax Program Installer"
    Write-Host ".. Os que mostram NL no nome precisam do Logoff para configurar o atalho"
    Write-Host "..  0 - Voltar para a tela inicial"
    Write-Host "..  1 - ARISP - NL"    
    Write-Host "..  2 - CAGED" 
    Write-Host "..  3 - Atualizar DAPI"    
    Write-Host "..  4 - DCTF - NL"
    Write-Host "..  5 - DECLAN - NL"    
    Write-Host "..  6 - DIEF ES" 
    Write-Host "..  7 - DIEF PA" 
    Write-Host "..  8 - DIEF PI"
    Write-Host "..  9 - DIRF 2018"
    Write-Host ".. 10 - DOT - NL"
    Write-Host ".. 11 - eDoc"
    Write-Host ".. 12 - GIA RJ ICMS - NL"
    Write-Host ".. 13 - GIA RS"
    Write-Host ".. 14 - IBGE"
    Write-Host ".. 15 - IRPF"
    Write-Host ".. 16 - Livro Eletronico"
    Write-Host ".. 17 - Perdcomp - NL"
    Write-Host ".. 18 - ReceitaNet - NL" 
    Write-Host ".. 19 - ReceitaNetBX - NL"
    Write-Host ".. 20 - SEF"
    Write-Host ".. 21 - Sefaz AL - Audora Desktop Tools"
    Write-Host ".. 22 - Sefaz PE"
    Write-Host ".. 23 - Sefaz PI - e-AGEAT"
    Write-Host ".. 24 - SEFIP"
    Write-Host ".. 25 - SVA"
    Write-Host ".. 26 - TED"
    Write-Host ".. 27 - TEDSEF"
    Write-Host ".. 28 - Validador LFE"
    Write-Host ".. 29 - Lacuna Web"
    Write-Host ".. 30 - CC Token"
    Write-Host ".. 31 - GIA ST"
    Write-Host ".. 32 - GIAM"
    Write-Host ".. 33 - Token SafeNet - eToken Pro (Aladdin), iKey 2032, Token 5100, Token 5110"
    Write-Host ".. 34 - DIRF 2015"
    Write-Host ".. 35 - DIRF 2016"
    Write-Host ".. 36 - DIRF 2017"
    $op = read-host ".Informe o numero da sua escolha"
     if ($op -eq "0") 
    {
        INI
    } 
     if ($op -eq "1") 
    {
        $FiscalName = "ARISP"
    }
    if ($op -eq "2") 
    {
        $FiscalName = "CAGED"
    }
    if ($op -eq "3") 
    {
        $FiscalName = "DAPI"
    }
    if ($op -eq "4") 
    {
        $FiscalName = "DCTF"
    }
    if ($op -eq "5") 
    {
        $FiscalName = "DECLAN"
    }
    
    if ($op -eq "6") 
    {
        $FiscalName = "DIEFES"
    }
    if ($op -eq "7") 
    {
        $FiscalName = "DIEFPA"
    }
    if ($op -eq "8") 
    {
        $FiscalName = "DIEFPI"
    }
    if ($op -eq "9") 
    {
        $FiscalName = "DIRF"
    }
    if ($op -eq "10") 
    {
        $FiscalName = "DOT"
    }
    if ($op -eq "11") 
    {
        $FiscalName = "eDoc"
    }
    if ($op -eq "12") 
    {
        $FiscalName = "GIAICMS"
    }
    if ($op -eq "13") 
    {
        $FiscalName = "GIARS"
    }
    if ($op -eq "14") 
    {
        $FiscalName = "IBGE"
    }
    if ($op -eq "15") 
    {
        $FiscalName = "IRPF"
    }
    if ($op -eq "16") 
    {
        $FiscalName = "Validador"
    }
    if ($op -eq "17") 
    {
        $FiscalName = "Perdcomp"
    }
    if ($op -eq "18") 
    {
        $FiscalName = "ReceitaNet" 
    }
    if ($op -eq "19") 
    {
        $FiscalName = "ReceitaNetBX"
    }
    if ($op -eq "20") 
    {
        $FiscalName = "SEF"
    }
    if ($op -eq "21") 
    {
        $FiscalName = "SefazAL" 
    }
    if ($op -eq "22") 
    {
        $FiscalName = "SefazPE" 
    }
    if ($op -eq "23") 
    {
        $FiscalName = "SefazPI"
    }
    if ($op -eq "24") 
    {
        $FiscalName = "SEFIP"
    }
    if ($op -eq "25") 
    {
        $FiscalName = "SVA" 
    }
    if ($op -eq "26") 
    {
        $FiscalName = "TED"
    }
    if ($op -eq "27") 
    {
        $FiscalName = "TEDSEF"
    }
    if ($op -eq "28") 
    {
        $FiscalName = "LFE"
    }
    if ($op -eq "29") 
    {
        $FiscalName = "LacunaWeb"
    }
    if ($op -eq "30") 
    {
        $FiscalName = "CCToken"
    }
    if ($op -eq "31") 
    {
        $FiscalName = "GIAST"
    }
    if ($op -eq "32") 
    {
        $FiscalName = "GIAM"
    }
    if ($op -eq "33") 
    {
        $FiscalName = "SafeNet"
    }
    if ($op -eq "34") 
    {
        $FiscalName = "DIRF15"
    }
    if ($op -eq "35") 
    {
        $FiscalName = "DIRF16"
    }
    if ($op -eq "36") 
    {
        $FiscalName = "DIRF17"
    }
    Remove-Item -Path "$Installpath\Config\fiscal.name" # Remove o arquivo de configuração
    Add-Content -Path "$Installpath\Config\fiscal.name" -Value "$FiscalName\$FiscalName.ps1"  # Cria um novo arquivo de configuração
    Start-Process -FilePath "$Installpath\Install\Remote.bat"  -WindowStyle Hidden # Inicia a instalação
    cls
    Write-Host ".. Todos os dados das instalacoes estao disponiveis nos LOGs em: $Installpath\LOG\$hostname.$FiscalName.log"
    pause
    INI # Volta para o menu
}

# Starts the program
inicio