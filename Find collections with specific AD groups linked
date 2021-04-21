######################################################
# Find collections with specific AD groups linked
######################################################

$collections = get-cmusercollection
foreach ($item in $collections){

if (Get-CMCollectionQueryMembershipRule -CollectionId $($item.collectionid)|where {
$_.queryexpression -match "serv_" -or
$_.queryexpression -match "NUANCE_" -or
$_.queryexpression -match "O365Licensing" -or
$_.queryexpression -match "SNOW_" -or
$_.queryexpression -match "Windows10_" -or
$_.queryexpression -match "Windows7_" -or
$_.queryexpression -match "CRM_" -or
$_.queryexpression -match "Desktops_" -or
$_.queryexpression -match "Exclusion_"
}){
write-host "$($item.collectionid): $($item.name): ($item.QueryExpression)"
}


}

######################################################
# Last updated 2020.05.22
######################################################
