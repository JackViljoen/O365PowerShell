# Connect to the SharePoint site
$siteUrl = "https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS"
$credentials = Get-Credential
$cnx = Connect-PnPOnline -Url $siteUrl -Credentials $credentials -ReturnConnection

# Specify the list name and fields
$listName = "_InvoiceArchive"

$List = Get-PnPList -Identity $listName -Connection $cnx

$Fields = Get-PnPField -List $listName -Connection $cnx
# $Field = Get-PnPField -List $listName -Identity "InvoiceDept" -Connection $cnx

$InterstingFields = @("Lookup", "User", "Choice")

foreach ($Field in $Fields) {
    if ($Field.TypeAsString -in $InterstingFields){
        #if ($Field.Title.IndexOf(":") -ne -1){
            $iName = $Field.InternalName
            $iTitle = $Field.Title
            Write-Host $Field.TypeAsString $Field.InternalName $Field.Title $Field.Id 
            #Remove-PnPField -Connection $cnx -List $listName -Identity $Field.Id -Force
            #Add-PnPField -Connection $cnx -List $listName -InternalName $iName -DisplayName $iTitle -Type Text
        #}
    }
}




