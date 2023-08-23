Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "") 

##=====================================================================================================================================
# Where is the Output Queue TAKE CARE !!!
# $PSFileSource = "C:\temp\InvoiceFiles\DAPS" ##LOCALTEST
# $PSFileSource = "\\plre.pacificlife.net\plfuat2\InvoiceFiles\DAPS" ##UAT
# $PSFileSource = "\\plre.pacificlife.net\plfprd2\InvoiceFiles\DAPS" ##PRODUCTION


##=====================================================================================================================================
# Site detail including the Time Window - an hour to run the thing in 
$Sites = @(
                @{
                    "active"     = $true
                    "siteUrl"    = "https://pacificlife.sharepoint.com/sites/PLRe-UMeDAPS"
                    "TimeZone"   = "GMT Standard Time"
                    "Contact"    = ""
                    "FileSource" = "Y:\"
                    "BU"         = "UMe"
                    "BUName"     = "UWME"
                } ,
                @{
                    "active"     = $true
                    "siteUrl"    = "https://pacificlife.sharepoint.com/sites/PLRe-AUDAPS"
                    "TimeZone"   = "AUS Eastern Standard Time"
                    "Contact"    = ""
                    "FileSource" = "Y:\"
                    "BU"         = "AU"
                    "BUName"     = "AUST"
                },
                @{
                    "active"     = $true
                    "siteUrl"    = "https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS"
                    "TimeZone"   = "GMT Standard Time"
                    "Contact"    = ""
                    "FileSource" = "Y:\"
                    "BU"         = "DC"
                    "BUName"     = "EURO"
                }
                
           )


function Move-AllInBatch() { 
    Param(  [parameter(position = 0)] $Invoices, 
            [parameter(position = 1)] $Message, 
            [parameter(position = 1)] $stage 
    ) 

   forEach($Failure in $Invoices){
        $Values = @{
            "PSVoucherNo"         = "$(Get-date -Format "yy-MM-dd")"; 
            "PSReturnMessage"     = "$($Message)" ;
            "wfSubStage"          = $Stage;
            "_UserField2"         = "";
            "_wfStatusChangeDate" = Get-date -Format "yyyy-MM-ddTHH:mm:ssZ";
            "_wfLockTime"         = $null;
        }
        if ($Failure.Id -is [Int]){ 
            #need to check we are not Passing the command a nothing as that will mess up the entite list  
            $a = set-PnpListItem -Connection $SiteConnection -List "Invoices" -Identity $Failure.Id -Values $Values
        } else {
            logActivity -Indent 2 -Type "6 Error" -Message "ERROR because Item ID is not an integer " -ForegroundColor Red
        }
    }
}


#======================================================================================================================================
#
## LOG START (PREAMBLE) 
#
#======================================================================================================================================

#first set up up for Relative Addressing in the file
$JobName = $MyInvocation.MyCommand.Name.split(".")[0]      #Get the Job name 
$r = $MyInvocation.MyCommand.Source                        #set up up for Relative Addressing everything is under O365PowerShell.
$rt = $r.Substring(0,$r.IndexOf("\O365PowerShell\") + 15)  # the path is everything up to and including the /O365PowerShell
Set-Location $rt

#SET UP LOGGING FOR THIS MODULE - LINK TO THE Code so it has the SCRIPT scope
.".\2-UTILITIES\SPLogger.ps1"
.".\2-UTILITIES\Utilities.ps1"
.".\3-CRON\CronSimple.ps1"

# Where is the site and the library to store the LOG into 
$LogSiteURL     = "https://pacificlife.sharepoint.com/sites/PLRe"
$LogLibName     = "wfHistoryEvents"

# Who are we savign the log as and conenct to the log site as them 
$logaccountName = "svc_sp_sync@Pacificlifere.com" 
$logencrypted   = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$logcredential  = New-Object System.Management.Automation.PsCredential($logaccountName, $logencrypted)
$LogConnection  = Connect-PnPOnline -Url $LogSiteURL -Credentials $logcredential -ReturnConnection  #-UseWebLogin #-Credentials $logcredential -ReturnConnection -TransformationOnPrem


# Set up the Logging control data (static)
$Script:LogControl.JobName        = $JobName
$Script:LogControl.LogLevel       = "1 Success";
$Script:LogControl.LogLib         = $LogLibName;
$Script:LogControl.LogConnection  = $LogConnection;
$Script:LogControl.LogContact     = "tim.ellidge@Pacificlifere.com";

# TYPES ARE : "1 Success", "2 Info", "3 Info", "4 Action", "5 Warning", "6 Error"
#Start the log with a simple entry
logActivity -Indent 0 -Type "1 Success" -Message "new root = $rt"
#======================================================================================================================================
#
# LOG SETUP END
#
#======================================================================================================================================
#


# SET UP THE CREDENTIALS TO USE TO GET TO THE PEOPLESOFT FILE LOCATION
#
$FUsername = "PLRE\svc_SP_Admin"
$FEncrypted = Get-Content ".\8-VAULT\3bdf7355-25fe-4df9-bb32-2a3d4b9b5874.txt" | ConvertTo-SecureString
$FileCreds = New-Object System.Management.Automation.PsCredential($FUsername, $FEncrypted)

$Source = "\\plreukppsdb01v\plfprd2\InvoiceFiles\DAPS"

New-PSDrive -Name Y -PSProvider FileSystem -Root $Source -Credential $FileCreds -ErrorAction SilentlyContinue -ErrorVariable ErrVar

if ($ErrVar.Count -gt 0) {
    write-host "Y drive already connected  $($ErrVar[0])"
} else {
    write-host "Added Y drive " -ForegroundColor Yellow
}



$PingPrefix = ""

#who are we doign the work as ?
$accountName = "svc_sp_sync@Pacificlifere.com" 
$encrypted   = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$credential  = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

foreach ($SI in $Sites){
    if ($SI.active){

        logActivity -Indent 2 -Type "2 Info" -Message "OK so it may run :-)" -ForegroundColor Green
        $PSFileSource = $SI.FileSource  
        #
        # POP THE CALL TO THE PAYLOAD IN HERE 
        #
       
        $files = get-ChildItem -Path "$($PSFileSource)\Output" | Where-Object {$_.Name -like "DAPS_$($SI.BUName)*" } | Sort-Object LastWriteTime -descending | Select-Object -First 1
        if($files.Count -eq 0){
            logActivity -Indent 2 -Type "2 Info" -Message "Nothing in the Queue matching the pattern DAPS_$($SI.BUName)* - will try again in an hour" -ForegroundColor Magenta 
            $BatchName = ""
        } else {
            if ($files.Count -eq 1){
                $PingPrefix += "$($SI.BU)_"
                           
                $Instance  = $files.Name.split("_")[1]
                $BatchName = $files.Name.split(".")[0].replace("_Output","")
                        
                $Data      = [IO.File]::ReadAllText($files.FullName)
                $Data      = $Data.replace(" ","~") # REMOVE SPACES
                $lines     = $Data.Split([Environment]::none) # turn it onto an array of seperate lines

                logActivity -Indent 2 -Type "4 Action" -Message "processing file:$($BatchName) for $($Instance) $($SI.siteURL)"

                #LETS CONNECT TO THE SITE AS WE WILL BE UPDATING RECORDS 
                $SiteConnection = $null
                $SiteConnection = Connect-PnPOnline -Url $SI.siteURL -Credentials $credential -ReturnConnection 
  
                $TheseInvoices  = Get-PnPListItem -Connection $SiteConnection -List "Invoices" | Where-object {$_.FieldValues._UserField2 -eq $BatchName -and $_.FieldValues.wfSubStage -eq "5.1 Pending Peoplesoft Coding"}

                if($TheseInvoices){ # DID WE FIND ANY MATCHING THE FILE ? 
                    $l= 1 #Set up a counter
                    foreach($LI in $lines){
                        if($LI.Length -gt 0){ # there can be empty CRLFs in the file usually at the end so ignore them 
                            $LI = $LI.replace("~"," ")  # Restore the spaces as they may be needed in the data 
                            $Fragments = $LI.split(",").replace('"','') # Now we chop up the line as it is comma delimited

                            $Response = @{                              # Probably not needed but lets make up an object with the line data that way its easier to query the response
                                "FileName"      = $Fragments[0];
                                "FileDate"      = get-date($Fragments[1]) -format "yy-MM-dd";
                                "DAPSInstance"  = $SI.BU;
                                "DAPSBusiness"  = $Fragments[3];
                                "DAPSInvoice"   = $Fragments[4];
                                "DAPSInvoiceDt" = $Fragments[5];
                                "PSID"          = $Fragments[6];
                                "PSCode"        = $Fragments[7];
                                "PSResponse"    = $Fragments[8];
                            }
                            logActivity -Indent 2 -Type "2 Info" -Message $LI
                            $l++
                            If ($Response.PSID -gt ""){
                                ## WOOHOO ITS GOT A VOUCHER NUMBER
                                logActivity -Indent 2 -Type "4 Action" -Message "Invoice:$($Response.DAPSInvoice) PS Voucher No:$($Response.PSID) is moving to stage 6 - $($Response.PSCode) $($Response.PSResponse)" -ForegroundColor Green
                                $goodOne = $TheseInvoices  | Where-object {$_.FieldValues.RefNo -eq $Response.DAPSInvoice}  ## -and $_.FieldValues.RAGDate.toString("yyyy-MM-dd") -eq (get-date($Response.DAPSInvoiceDt)).toString("yyyy-MM-dd")}
                                
                                if($goodOne){
                                    $Values = @{
                                        "PSVoucherNo"         = "$($Response.FileDate) : $($Response.PSID)"; 
                                        "PSReturnMessage"     = "$($Response.PSCode) - $($Response.PSResponse)";  
                                        "wfSubStage"          = "6.0 Pending approval"; 
                                        "_wfStatusChangeDate" = $LoopStart;
                                        "_wfLockTime"         = $null;
                                    }
                                    $a = set-PnpListItem -Connection $SiteConnection -List "Invoices" -Identity $goodOne.Id -Values $Values
                                    #TODO TEST ON UPDATE 
                                } else {
                                    logActivity -Indent 2 -Type "2 Info" -Message "Invoice:$($Response.DAPSInvoice) was not found to update... " -ForegroundColor red
                                }
                            } else {
                                if($Response.PSCode -ne "BAD" -and $Response.PSCode -ne "LCN"){
                                    ## It cant be processed
                                    logActivity -Indent 2 -Type "4 Action" -Message "Invoice:$($Response.DAPSInvoice) PS Voucher No:$($Response.PSID) is moving to stage 5.2 - $($Response.PSCode) $($Response.PSResponse)" -ForegroundColor Green
                                    #This is a terrible match is the Invoice Numbers are the same 
                                    $goodOne = $TheseInvoices  | Where-object {$_.FieldValues.RefNo -eq $Response.DAPSInvoice} | Select-Object -First 1
                                    if($goodOne){
                                        $Values = @{
                                                "PSVoucherNo"         = "$($Response.FileDate) : $($Response.PSID)"; 
                                                "PSReturnMessage"     = "$($Response.PSCode) - $($Response.PSResponse)";  
                                                "wfSubStage"          = "5.2 Peoplesoft Coding Issue"; 
                                                "_wfStatusChangeDate" = $LoopStart;
                                        }
                                        $a = set-PnpListItem -Connection $SiteConnection -List "Invoices" -Identity $goodOne.Id -Values $Values
                                        ## TO DO TEST ACTUAL UPDATE 
                                        logActivity -Indent 2 -Type "4 Action" -Message "Invoice:$($Response.DAPSInvoice) Response Code:$($Response.PSCode) is going to stage 5.2 - $($Response.PSResponse)" -ForegroundColor Yellow
                                    } else {
                                        logActivity -Indent 2 -Type "5 Warning" -Message "Invoice:$($Response.DAPSInvoice) was not found to update... " -ForegroundColor red
                                    }
                                } else {
                                    logActivity -Indent 2 -Type "7 Alert" -Message "ALL $($TheseInvoice.count) Invoices in :$($Response.FileName) are going to 5.2 there has been a File error Helpdesk ticket raised" -ForegroundColor RED
                                    #DON'T need to test here as we already tested 
                                    Move-AllInBatch -Invoices $TheseInvoices -Message "PS Return file failure - Contact PeopleSoft Support" -Stage "5.2 Peoplesoft Coding Issue"; 
                                } 
                            }
                        }
                    }
                } else {
                    logActivity -Indent 2 -Type "6 Error" -Message "ERROR WE couldn't find ANY invoices in $($SI.siteUrl) matching the PS submission of $($BatchName)" -ForegroundColor red
                }

                # WE have finished processing it for good or ill so move it to the archive :-)
                $NewFileName = $files.Name.replace(".csv", "_Output.csv") 
                copy-Item $files.FullName "$($PSFileSource)\Archive\$($NewFileName)"
                copy-Item $files.FullName "C:\temp\InvoiceFiles\DAPS\Archive\$($NewFileName)" #TEMP LOCAL STORAGE SO I CAN GET TO IT WHILST TESTING AND MONITORING  
                remove-Item $files.FullName #GET RID FROM THE OUTPUT QUEUE AS ITS NOW PROCESSED
            } else {
                logActivity -Indent 2 -Type "2 Info" -Message "We have $($files.Count) PS Files in the Queue for $($SI.BUName) THIS IS WRONG!" -ForegroundColor Magenta 
            }
        } 
    } else {
        logActivity -Indent 2 -Type "2 Info" -Message "$($SI.siteUrl) is Inactive" -ForegroundColor magenta
    }
    LogActivity -Indent 0 -Type "2 Info" -Message "JOB Finished for $($SI.siteUrl)" -logAction "WriteQuiet"
}


#======================================================================================================================================
#
# TASK END HERE - CLOSE OUT THE LOG AND REGISTER THE PING
#
#======================================================================================================================================
$JobDuration = ((get-date) -  $Script:LogControl.LogFirstCall).TotalSeconds # how long did the processing take
$ping        = @{"LastAlive" = get-date ; "Duration(s)" = $JobDuration ;}

#check if its the first one if so create a directory otherwise just save it  
if((test-path -Path ".\9-PINGS\$($JobName)") -eq $false){New-Item -Path ".\9-PINGS\$($JobName)" -ItemType directory}
$a           = $ping | Out-File -FilePath ".\9-PINGS\$($JobName)\$($PingPrefix)$($JobName)$(Get-date -Format "yyMMdd-HH_mm").txt" 

#WriteQuiet - only write a record IF the max error exceeds the limit $Script:LogControl.LogLevel 
#Write always writes at least one line  
#logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished" -logAction "WriteQuiet"
#>


# Lets get a local copy of the files


# $PSFileSource = "C:\temp\InvoiceFiles\DAPS" ##LOCALTEST
# $PSFileSource = "\\plre.pacificlife.net\plfuat2\InvoiceFiles\DAPS" ##UAT
# $PSFileSource = "\\plre.pacificlife.net\plfprd2\InvoiceFiles\DAPS" ##PRODUCTION
# get-ChildItem -Path "$($PSFileSource)" -Recurse | Sort-Object LastWriteTime -Descending

#RECOVER TEH PROD PROCESSED FILES FOR ANAYSIS
#copy-Item  "\\plre.pacificlife.net\plfprd2\InvoiceFiles\DAPS\*" "C:\temp\InvoiceFiles\DAPS\" -Force #TEMP LOCAL STORAGE SO I CAN GET TO IT WHILST TESTING AND MONITORING  
#copy-Item  "\\plre.pacificlife.net\plfprd2\InvoiceFiles\DAPS\Output\*" "C:\temp\InvoiceFiles\DAPS\Output\" -Force #TEMP LOCAL STORAGE SO I CAN GET TO IT WHILST TESTING AND MONITORING  
#copy-Item  "\\plre.pacificlife.net\plfprd2\InvoiceFiles\DAPS\Error\*" "C:\temp\InvoiceFiles\DAPS\Error\"  -Force
#copy-Item  "\\plre.pacificlife.net\plfprd2\InvoiceFiles\DAPS\Archive\*" "C:\temp\InvoiceFiles\DAPS\Archive\"  -Force