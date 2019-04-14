# Script para correção da falha conhecida da 1709 
# AUTOR  : Ivo Dias
# VERSAO : 1.0.MSC

#Limpa a tela ao iniciar o programa
cls

#Deleta a chave de registro responsavel pela migração dos programas da loja
reg delete “HKCU\Software\Microsoft\Windows NT\CurrentVersion\TileDataModel\Migration\TileStore” /va /f
#Recria as aplicações padrões
get-appxpackage -packageType bundle |% {add-appxpackage -register -disabledevelopmentmode ($_.installlocation + "\appxmetadata\appxbundlemanifest.xml")} $bundlefamilies = (get-appxpackage -packagetype Bundle).packagefamilyname
get-appxpackage -packagetype main |? {-not ($bundlefamilies -contains $_.packagefamilyname)} |% {add-appxpackage -register -disabledevelopmentmode ($_.installlocation + "\appxmanifest.xml")}

 #Escreve na tela informando que encerrou
 Write-Host O procedimento foi completado com sucesso
 Pause