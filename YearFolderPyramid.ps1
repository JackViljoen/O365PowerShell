# Connect
$SiteURL     = "https://pacificlife.sharepoint.com/sites/PLRe-tDivSPOServiceDesk"
$Credentials = Get-Credential
$cnx         = Connect-PnPOnline -Url $SiteURL -Credentials $Credentials
$ListName    = "FolderPyramid"

# Years
$Years    = @(2022, 2023, 2024)
$Quarters = @("Q1", "Q2", "Q3", "Q4")
$Months   = @("1-Jan", "2-Feb", "3-Mar", "4-Apr", "5-May", "6-Jun", "7-Jul", "8-Aug", "9-Sep", "10-Oct", "11-Nov", "12-Dec")

# Connect to list
$List = Get-PnPList -Identity $ListName -Connection $cnx

# Loop through Years
foreach ($Year in $Years) {
    $YearFolder = $Year.ToString()
    $YearFolder = Add-PnPFolder -Name $YearFolder -Folder $List.RootFolder
    
    # Loop through Quarters
    foreach ($Quarter in $Quarters) {
        $QuarterFolder = Add-PnPFolder -Name $Quarter -Folder $YearFolder

        # Loop through Months
        foreach ($Month in $Months) {
            $MonthFolder = Add-PnPFolder -Name $Month -Folder $QuarterFolder

        }
        
    }
}

Disconnect-PnPOnline

