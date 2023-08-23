Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "") 

function writeoutbatch($batch){
    #log out the batches for DEBUG purposes 
    write-host ($batch[0].StageName) -ForegroundColor Yellow
    foreach ($BA in $batch){ 
        write-host ($BA| Format-Table | Out-String)
    }
}

## loop these for all of the sites :-) basically the site paramaters 
$DAPSInstances = @(
    @{
        "active"        = $true
        "path"          = "https://pacificlife.sharepoint.com/sites/PLRe-AUDAPS"
        "accountsEmail" = "Accounts.PayableAUS@PacificLifere.com"
        "keyGroup"      = "AUDAPS Releasers"
        "EmailDebug"    = $false
        "BU"            = "AU"
    }
    @{
        "active"        = $true
        "path"          = "https://pacificlife.sharepoint.com/sites/PLRe-UMeDAPS"
        "accountsEmail" = "accounts@underwriteme.co.uk"
        "keyGroup"      = "UMEDAPS Releasers"
        "EmailDebug"    = $false
        "BU"            = "UME"
    },
    @{
        "active"        = $true
        "path"          = "https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS"
        "accountsEmail" = "accountspayableUK@PacificLifere.com"
        "keyGroup"      = "DC DAPS Releasers"
        "EmailDebug"    = $false
        "BU"            = "DC"
    },
        @{
        "active"        = $false
        "path"          = "https://pacificlife.sharepoint.com/sites/PLRe-REDAPS"
        "accountsEmail" = "tim.ellidge@PacificLifere.com"
        "keyGroup"      = "DCDAPS Releasers"
        "EmailDebug"    = $true
        "BU"            = "RE"
    }
)

#=============================================================================================================================
# EMAIL DETAIL BLOCK AND HYDRATE FUNCTION 
#=============================================================================================================================

$BatchEmails = @{
   "8.0_Manual" = @{
        "Subject" = "DO NOT REPLY: Batch Release - Manual Batch or payment to release - [BatchName]"
        "Message" = "Hi [FirstName] <br/><br/>"`
            + "Invoice Release Request - Batch: [BatchName]  contains: [Count] invoice(s)  Total value [Value]<br/>"`
            + "<br/>Please click the below link to get to sharepoint. <br/>"`
            + "<a href='[path]/SitePages/Reconcilliation.aspx?batch=[BatchName]&batchtype=Manual&batchstage=0'>"`
            + " Release / reconcilliation page</a><p>Note: After you have finished processing the invoices in this batch the link will no longer work. </p>Thanks in advance,<br/><br/>DAPS."`
            + "<br/><b>To contact your accounts team click <a href='mailto:[accountsEmail]'>here</a></b>"
    }
   "8.1_BACS_Matched" = @{
        "Subject" = "DO NOT REPLY: Batch Release BANK file uploaded & Invoices matched - Please release batch [BatchName]"
        "Message" = "Hi [FirstName] <br/><br/>"`
            + "Invoice Release Request - Batch: [BatchName]  contains: [Count] invoice(s)  Total value [Value]<br/>"`
            + "<br/>Please click the below link to get to sharepoint. <br/>"`
            + "<a href='[path]/SitePages/Reconcilliation.aspx?batch=[BatchName]&batchtype=BACS&batchstage=1'>"`
            + " Release / reconcilliation page</a><p>Note: After you have finished processing the invoices in this batch the link will no longer work. </p>Thanks in advance,<br/><br/>DAPS."`
            + "<br/><b>To contact your accounts team click <a href='mailto:[accountsEmail]'>here</a></b>"
    }
    "8.2_BACS_Reject" = @{
        "Subject" = "DO NOT REPLY: Batch Release - PLEASE BANK  and forward on"
        "Message" = "Hi Accounts Payable, <br/><br/>"`
            + "Items have been removed from the ACH file. (see below)"`
            + "<br/>The link below will direct you to the expenses already approved, please confirm the bank details on the BACS line match to the ACH file. "`
            + "<br/></br><a href='[path]/SitePages/Reconcilliation.aspx?batch=[BatchName]&batchtype=[BatchType]&batchstage=2'> Release / Reconcilliation page</a>"`
            + "</p>Thanks in advance,<br/><br/>DAPS."
        }

    "8.3_Secondary" = @{
        "Subject" = "DO NOT REPLY: Batch Release Secondary release required."
        "Message" = "Hi [SecondaryFirstName] <br/><br/>"`
            + "Secondary Release Request - Batch: [BatchName]  contains: [Count] invoice(s)  Total value [Value]<br/>"`
            + "<br/>Please click the below link to get to DAPS. <br/>"`
            + "<a href='[path]/SitePages/Reconcilliation.aspx?batch=[BatchName]&batchstage=3'>"`
            + "Secondary releaser page</a><p>Note: After you have finished processing the invoices in this batch the link will no longer work. </p>Thanks in advance,<br/><br/>DAPS."`
            + "<br/><b>To contact your accounts team click <a href='mailto:[accountsEmail]'>here</a></b>"
    }
    "8.3_Secondary_Reject" = @{                
        "Subject" = "DO NOT REPLY: Batch Release [BatchName] Secondary release not complete."
        "Message" = "Hi  Accounts Payable, <br/><br/>"`
            + "Secondary releaser [Secondary] has rejected one or more of the items in this batch '[BatchName]', please contact them for details. <br/>"`
            + "other items in this batch '[BatchName]', have been moved back to 8.1 please contact them for details. <br/>"`
            + "</p>Thanks in advance,<br/><br/>DAPS." 
     }
}

function hydrateEmail(){
    Param(
        [parameter(position = 0)] $BatchStage,
        [parameter(position = 1)] $BatchDetail,
        [parameter(position = 2)] $AccountsEmail, 
        [parameter(position = 2)] $sitepath 
    )
    #prepare the low hanging substitutions. BEWARE WE NEED THE VALUES NOT THE OBJECT
    $responseSubject  = [string] $BatchEmails[$BatchStage].Subject
    $responseMessage  = [string] $BatchEmails[$BatchStage].Message
   
    $firstname        =  $BatchDetail.Primary.split('.')[0]
    $secondaryfirstname  = if($BatchDetail.Secondary) {$BatchDetail.Secondary.split('.')[0]}
    $responseMessage = $responseMessage.Replace("[FirstName]", $firstname)
    $responseMessage = $responseMessage.Replace("[SecondaryFirstName]", $secondaryfirstname)
    $responseMessage = $responseMessage.Replace("[accountsEmail]", $AccountsEmail)
    $responseMessage = $responseMessage.Replace("[path]", $sitepath)
    foreach($k in $BatchDetail.Keys){
        $responseSubject = $responseSubject.Replace("[$($k)]",$BatchDetail[$k] )
        $responseMessage = $responseMessage.Replace("[$($k)]",$BatchDetail[$k] )
    }
    write-host $responseSubject -ForegroundColor Green
    write-host $responseMessage 
    return @{"Subject" = $responseSubject
             "Message" = $responseMessage
    }
            
}

#=============================================================================================================================
# END EMAIL CONTENT BIT 
#=============================================================================================================================

function BatchStatus() {
    Param(
        [parameter(position = 0)] $Invoices,
        [parameter(position = 1)] $St,
        [parameter(position = 2)] $Bt
    )
    write-host "working on $Bt [$St]" -ForegroundColor Cyan
    $Rows = $Invoices | Where-Object { $_.FieldValues._PaymentBatch -eq $Bt }  | Sort-Object { $_.FieldValues.InvoiceAmount }
    $Result = @{
            BatchName    = $Bt
            StageName    = $St
            Count        = $Rows.Count
            Value        = 0
            AtStage      = 0
            Under        = 0
            Over         = 0
            Rejected     = 0
            NotProcessed = 0
            Failed       = 0
            Earliest     = get-date("2099-12-31")
            Latest       = get-date("2000-01-01")
            AgeMin       = 0
            BatchType    = ""
            Primary      = ""
            Secondary    = ""
            SendEmail    = $false
    }
    if($Rows.count -eq 1){
        #A LOT IS GOT FROM THE ONLY RECORD IN THE BATCH
        $Result.BatchType = $Rows.FieldValues._PaymentType
        $Result.Primary   = $Rows.FieldValues.AssignedTo1.Email
        $Result.Secondary = $Rows.FieldValues.BusinessOwner.Email
        $Result.SendEmail = $Rows.FieldValues.EmailAlert
    } else {
        #A LOT IS GOT FROM THE FIRST RECORD IN THE BATCH
        $Result.BatchType = $Rows[0].FieldValues._PaymentType
        $Result.Primary   = $Rows[0].FieldValues.AssignedTo1.Email
        $Result.Secondary = $Rows[0].FieldValues.BusinessOwner.Email
        $Result.SendEmail = $Rows[0].FieldValues.EmailAlert
    }
    
    foreach ($Ro in $Rows) {
        #part 1 over or under at stage
        if ($Ro.FieldValues.wfSubStage -gt $St) {
            $Result.Over++
        }
        else {
            if ($Ro.FieldValues.wfSubStage -lt $St) {
                $Result.Under++
            }
            else {
                $Result.AtStage++
                $Result.Value += $Ro.FieldValues.InvoiceAmount
            }
        }
        #part 1 a 
        if ($Ro.FieldValues.wfSubStage -lt "2.0") {
                $Result.Rejected++
        }

        #part 2 Matched or failed 
        if ($Ro.FieldValues._BACSREF -eq $null -or $Ro.FieldValues._BACSREF -eq "") {
            $Result.NotProcessed++
        }
        else {
            if ($Ro.FieldValues._BACSREF.IndexOf("FAILED") -gt -1) {
                $Result.Failed++
            }
        }
        #part 3 get the times if there are any 
        if($Ro.FieldValues._SystemUpdateDate){ 
            $Result.Earliest = if ($Ro.FieldValues._SystemUpdateDate -lt $Result.Earliest) {$Ro.FieldValues._SystemUpdateDate}
            $Result.Latest   = if ($Ro.FieldValues._SystemUpdateDate -gt $Result.Latest)   {$Ro.FieldValues._SystemUpdateDate}
        }
    }
    #$ts = New-TimeSpan -Start $Result.Latest -End (get-date)
    #$Result.AgeMin = $ts.TotalMinutes   # how long since the last match
   
    return  $Result
}

function ListInvoiceItems() {
    Param(
        [parameter(position = 0)] $List,
        [parameter(position = 1)] $St,
        [parameter(position = 2)] $Bt
    )
    $ReturnStr = ""
    $Rows = $List.items | Where-Object { $_.FieldValues.wfSubStage -eq $St -and $_.FieldValues._PaymentBatch -eq $Bt }  | Sort-Object { $_.FieldValues._SplitAmount }
    foreach ($Row in $Rows) {
        $ReturnStr += "    $($Row['Name']) for ($($Row['Currency'])) $($Row['_SplitAmount']) Note: $($Row['StageNote'])<br/>"
    }
    return $ReturnStr
}

function BatchSetEmailFlag () {
    Param(
        [parameter(position = 1)] $BatchName,
        [parameter(position = 2)] $SendMail
    )

    logActivity -Indent 0 -Type "2 Info" -Message  "Setting Email to $($SendMail) for all invoices in the batch  $($BatchName)"
    $Rows = Get-PnPListItem -Connection $Connection -List "Invoices" | Where-Object { $_.FieldValues._PaymentBatch -eq $BatchName } 
    foreach ($Row in $Rows) {
        $a = set-PnpListItem -Connection $Connection -List "Invoices" -Identity $Row.Id -Values @{"EmailAlert" = $SendMail} 
    }
}

function UpdateItemsTo() {
    # used to send things to stage 9 either as a pre release background task or as a DD approved
    # if the comment is $null dotn add a history record 
    Param(
        [parameter(position = 0)] $Items,
        [parameter(position = 1)] $toStage,
        [parameter(position = 2)] $Action,
        [parameter(position = 4)] $Comment = $null, # dont log i no comment 
        [parameter(position = 5)] $Who = $null,
        [parameter(position = 6)] $EmailAlert = $false # By default 
    )


    foreach ($thing in $Items) {
        $fromStage = $thing.FieldValues.wfSubStage
        if ($fromStage -ne $toStage) {
            $timespan   = if($thing.FieldValues._wfStatusChangeDate) {New-TimeSpan -Start $($thing.FieldValues._wfStatusChangeDate) -End  (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } else {0}
            $updateValues = @{
                "EmailAlert"          = $EmailAlert
                "wfSubStage"          = $toStage
                "_wfLockTime"         = $null
                "_wfStatusChangeDate" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            $a = Set-PnPListItem -Connection $Connection -List "Invoices" -Identity $thing.Id -Values $updateValues

            if ($Comment -ne $null) {
                if ($Who -eq $null) {
                    $Who = $thing.FieldValues.Editor.Email  # if we didn't sent a pretend user then use the last changed user
                }
                $folderName = "F$($thing.Id)"

                $HistoryRecord = @{
                    "Title"           = [string] $thing.Id + "-" + $Who
                    "_wfFormID"       = $thing.Id
                    "_wfTime"         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "_wfUser"         = $Who
                    "_wfFormType"     = $thing.FieldValues._wfFormType
                    "_wfAction"       = $Action
                    "_wfStreamStatus" = "$toStage ||"
                    "_wfStageChange"  = $toStage
                    "_wfPrevStage"    = $fromStage
                    "_wfStreamTime0"  = $timespan.TotalHours
                    "_wfStreamTime1"  = 0
                    "_wfStreamTime2"  = 0
                    "UserLogData"     = '[{"Note":"' + $Comment + '"}]'
                }                      
                $newHist = Add-PnPListItem -Connection $Connection -List "Lists/wfHistoryF" -Folder "F$($thing.Id)" -Values $HistoryRecord
                logActivity -Indent 0 -Type "4 Action" -Message  "Created a history record item: $($thing.Id) moving it to stage :  $($toStage) on behalf of $($Who)"
            }
            else {
                logActivity -Indent 0 -Type "2 Info" -Message  "did not Create a history record item: $($thing.Id) just moved it from:$($SU.FieldValues.wfSubStage) to stage:$($toStage)"
            }
        }
        else {
            logActivity -Indent 0 -Type "2 Info" -Message  "nothing to do for $($thing.Id) -  $($thing.FieldValues.Title)"
        }
    }
}

#======================================================================================================================================
#
## LOG START (PREAMBLE) 
#
#======================================================================================================================================

$JobName = $MyInvocation.MyCommand.Name.split(".")[0]      #Get the Job name 
$r = $MyInvocation.MyCommand.Source                        #set up up for Relative Addressing everything is under O365PowerShell.
$rt = $r.Substring(0,$r.IndexOf("\O365PowerShell\") + 15)  # the path is everything up to and including the /O365PowerShell
Set-Location $rt

#SET UP LOGGING FOR THIS MODULE - LINK TO THE Code so it has the SCRIPT scope
.".\2-UTILITIES\SPLogger.ps1"
.".\2-UTILITIES\Utilities.ps1"

# LEVELS ARE : "1 Success", "2 Info", "3 Info", "4 Action", "5 Warning", "6 Error"
$Log = @{
    "SiteURL"     = "https://pacificlife.sharepoint.com/sites/PLRe"
    "LibName"     = "wfHistoryEvents"
    "AccountName" = "svc_sp_sync@Pacificlifere.com" 
    "Password"    = $ValidUsers[$env:USERNAME]
    "Contact"     = "tim.ellidge@Pacificlifere.com";
    "Level"       = "1 Success"; 
}

$L = start-Log -Log $Log -ID $ID -RuleName $JobName
#======================================================================================================================================
# LOG SETUP END
#=====================================================================================================================================

    $accountName = "svc_sp_sync@Pacificlifere.com" 
    $encrypted   = Get-Content $ValidUsers[$env:USERNAME] | ConvertTo-SecureString
    $credential  = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

    ## ALL LOOPABLE BELOW HERE ??? 
    foreach ($instance in $DAPSInstances) {
        if($instance.active){
            $path = $instance.path 
            $accountsEmail = $instance.accountsEmail
            $keyGroup = $instance.keyGroup

            $Connection  = Connect-PnPOnline -Url $path -Credentials $credential -ReturnConnection
            $MyPeople    = Get-PnPGroupMember  -Connection $Connection -Group $keyGroup
            logActivity -Indent 0 -Type "2 Info" -Message  "Info1 | BACS Job |  |  |  loaded site and lists and $($keyGroup) contains $($Mypeople.count) people - now seeing if we have any relavent invoices  "

##~=====================================================================================================================================================
#~ PROCESS BACS FILES AND ADD BACS RECORDS IF NEEDED
##~=====================================================================================================================================================

            $ActiveFiles = Get-PnPListItem -Connection $Connection -List "InvoiceBACSFiles"  | Where-Object { $_.FieldValues._SystemUpdateDate -eq $null }

            logActivity -Indent 0 -Type "2 Info" -Message  "We have $($ActiveFiles.count) Bank BACS files to process "

            foreach ($item in $ActiveFiles) {
                logActivity -Indent 0 -Type "2 Info" -Message  "Info1 | BACS Job | $($item.Id)  |   $($item.FieldValues.FileLeafRef) "
                $Data  = Get-PnPFile -Connection $Connection  -Url $item.FieldValues.FileRef -AsString
                $Data      = $Data.replace(" ","~") # REMOVE SPACES THIS WILL HELP THE LINE EXTRACTION 
                $Lines = $Data.Split([Environment]::none) # turn it onto an array of seperate lines - perfect  - the first 4 lines are shite we dont beed
                $iCount = 0

                for ($i = 8; $i -le ($Lines.length - 8); $i++) {
                    if($Lines[$i].Length -gt 0){ # there can be empty CRLFs in the file usually at the end so ignore them 
                        $Lines[$i]     = $Lines[$i].replace("~", " ") # Pop them back in again now we are out of the other site  
                        # add a row for each line
                        $iName = ($item.FieldValues.FileLeafRef).replace(".TXT","")
                        $LineValues = @{
                            "Title"           = $iName + "[" + ($i - 3) + "]"
                            "_BACSREF"        = $iName
                            "_PayeeSortCode"  = $Lines[$i].Substring(0, 6)  # to vary you can use [$i%9]
                            "_PayeeAccountNo" = $Lines[$i].Substring(6, 8)
                            "_SplitAmount"    = ( [int] $Lines[$i].Substring(35, 11)) / 100
                            "InvoiceNo"       = $Lines[$i].Substring(64, 17).Trim()
                            "BACSVendor"      = $Lines[$i].Substring(82, 18).ToUpper().Trim()
                        }
                        $iCount++
                        $a= Add-PnPListItem -Connection $Connection -List "Lists/InvoiceBACSItems" -Values $LineValues
                    } 
                }
                logActivity -Indent 0 -Type "4 Action" -Message  "Info1 | BACS Job |  |  |  added  $($iCount) BACS entries"
                $a = set-PnpListItem -Connection $Connection -List "InvoiceBACSFiles" -Identity $item.Id -Values @{"_SystemUpdateDate" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")}
            }

##~=====================================================================================================================================================
#~ NOW GET THE INVOICES IN RANGE FOR THE REST OF THE PROCESSING (DD - WIRE - MANUAL - AND BACS)
##~=====================================================================================================================================================
            $InterestignInvoices = Get-PnPListItem -Connection $Connection -List "Invoices" | Where-Object { $_.FieldValues.wfSubStage -gt "7.0" -and $_.FieldValues.wfSubStage -lt "9.0"  }

##~=====================================================================================================================================================
#~ FILTER THE DD ONES Direct debit Process approved items to 9 directly (see - do not batch - JUST release)
##~=====================================================================================================================================================

            $Invoices = $InterestignInvoices | Where-Object { $_.FieldValues.wfSubStage -eq "7.0 Ready for payment" -and $_.FieldValues._PaymentType -eq "Direct Debit" }
            if ($Invoices.length -eq 0) {
                logActivity -Indent 0 -Type "2 Info" -Message  "There are no DD invoices ready for payment"
            } else {
                logActivity -Indent 0 -Type "2 Info" -Message  " have $($Invoices.length) invoices that are approved with a payment type of Direct debit"
                UpdateItemsTo -Items  $Invoices -toStage "9.0 Released" -Action '<i class="fas fa-pound-sign fa-2x fa-fw iGreen"></i> Released' -Comment "Direct Debit invoices do not need release approval " -Who "Automatic" 
            }

##~ =====================================================================================================================================================
#~ BATCH NAME PREPARATION _ Split out the Users From the batch name but filter by the ones with a pipe in them NOW USING EMAILS :=) 
##~ =====================================================================================================================================================
        $Invoices = $InterestignInvoices | Where-Object { $_.FieldValues.wfSubStage -eq "8.0 Batched" -and $_.FieldValues._PaymentBatch.IndexOf("|") -gt 0 } | Sort-Object { $_.FieldValues._PaymentBatch } #that are at the right stage

            foreach ($Invoice in $Invoices) {
                # set some field data ie unpick it from the batch name but test if it needs doign this is why there are bacs and manual in here
                logActivity -Indent 0 -Type "4 Action" -Message  "InvoiceNo: $($Invoice.FieldValues.RefNo) payment batchname: $($Invoice.FieldValues._PaymentBatch) needs to be split up"
                $BatchNameBits = ($Invoice.FieldValues._PaymentBatch).split("|")
                #Get the email addreses out and remove any helpful HTML 
                $p1   = if($BatchNameBits[1]) {$BatchNameBits[1] -replace '<[^>]+>',''} else {$null}
                $p2   = if($BatchNameBits[2]) {$BatchNameBits[2] -replace '<[^>]+>',''} else {$null}

                #CHECK THE PEOPLE ARE IN THE RIGHT GROUP AND RAISE ISSUE IF NOT.  WE DO THIS CHECK AS THEY NEED ACCESS IN ADDITION TO THE EMAIL 
                $Assigned      = $MyPeople | Where-Object { $_.email.toLower() -eq $p1.trim().toLower() }
                $BusinessOwner = if($p2) {$MyPeople | Where-Object { $_.email.toLower() -eq $p2.trim().toLower()} } else {$null}

                if(!$Assigned -or ($p2 -and !$businessOwner)){
                    $ErrorMessage = "either $($p1) or $($p2) is not in the group $($keyGroup)"
                    logActivity -Indent 0 -Type "6 Error" -Message  $ErrorMessage
                    #$Mail = smtp-SendEmail -RecipientEmails $accountsEmail -MessageText $ErrorMessage -Subject "ERROR" --From $accountsEmail -BCC "tim.ellidge@pacificlifere.com" -EmailDebug $instance.EmailDebug
                }

                $BatchData = @{
                    "_PaymentBatch" = $BatchNameBits[0] 
                    "AssignedTo1"   = $p1
                    "BusinessOwner" = $p2
                }
                $u = Set-PnPListItem -Connection $Connection -List "Invoices" -Identity $Invoice.Id -Values $BatchData 
            }

##~=====================================================================================================================================================
#~ NOW JUST GET THE INVOICES IN RANGE FOR BACS MATCHING (BUT GET FROM THE SOURCE) 
##~=====================================================================================================================================================
            $Invoices = Get-PnPListItem -Connection $Connection -List "Invoices" | Where-Object { $_.FieldValues.wfSubStage -eq "8.0 Batched" -and $_.FieldValues._PaymentType -eq "BACS" } | Sort-Object { $_.FieldValues._PaymentBatch } #that are at the right stage
        
            if(!$Invoices){
                logActivity -Indent 0 -Type "2 Info" -Message  "No BACS invoices that are batched but not yet matched"
            } else {
                logActivity -Indent 0 -Type "2 Info" -Message  "We have $($Invoices.length) invoices that are batched but not yet matched"
                $UnmatchedBACSItems = Get-pnpListItem -Connection $Connection -List "InvoiceBACSItems"   | Where-Object { $_.FieldValues._SystemUpdateDate -eq $null -or $_.FieldValues._SystemUpdateDate -eq "" }
                logActivity -Indent 0 -Type "2 Info" -Message  "AND we have ($($UnmatchedBACSItems.length)) BACS items that are in HSBC but not yet matched..."
                if(!$UnmatchedBACSItems){
                    logActivity -Indent 0 -Type "2 Info" -Message  "No BACS ITEMS to be matched at this time"
                } else {
                    foreach ($Invoice in $Invoices) {
                        logActivity -Indent 0 -Type "2 Info" -Message  "MATCHING Looking for - InvoiceNo: $($Invoice.FieldValues.RefNo) amount: $($Invoice.FieldValues.InvoiceAmount) vendor: $($Invoice.FieldValues._Vendor)"
                        $MatchedBACSItem = $null # make the test easier and really  make sure the last loop didnt blur into this one
                        $MatchedBACSItem = $UnmatchedBACSItems | Where-Object { $_.FieldValues.InvoiceNo.Trim() -eq $Invoice.FieldValues.RefNo.Trim() -and $_.FieldValues._SplitAmount -eq $Invoice.FieldValues.InvoiceAmount -and $Invoice.FieldValues._Vendor.toUpper().IndexOf($_.FieldValues.BACSVendor) -ne -1 }
                        if ($MatchedBACSItem) {
                            logActivity -Indent 0 -Type "4 Action" -Message  "Info | BACS Job |  |  |  we have one :  $($MatchedBACSItem.FieldValues.BACSVendor) No:  $($MatchedBACSItem.FieldValues.InvoiceNo) ($($MatchedBACSItem.FieldValues.Currency)): $($MatchedBACSItem.FieldValues._SplitAmount) "
                            $invValues = @{
                                "_BACSREF"          = $MatchedBACSItem.Id;
                                "wfSubStage"        = "8.1 Matched";
                                "EmailAlert"        = $true;
                                "_SystemUpdateDate" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            }
                            # process both records by updating them
                            $a = set-PnPListItem -Connection $Connection -List "InvoiceBACSItems" -Identity $MatchedBACSItem.Id -Values @{"_SystemUpdateDate" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")}
                        } else {
                            logActivity -Indent 0 -Type "3 Info" -Message  "Error | BACS Job |  |  |  we don't have a BACS match for invoice: $($Invoice.FieldValues.Title) "
                            $invValues = @{"_BACSREF" = "FAILED to find a matching BACS Record for:($($Invoice.FieldValues._PaymentBatch))  at $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") : will try later"}
                        }
                        $a = set-PnPListItem -Connection $Connection -List "Invoices" -Identity $Invoice.Id -Values $invValues 
                    }
                }
            }

##~=====================================================================================================================================================
#~ NOW RE-GET THE INVOICES IN RANGE FOR THE REST OF THE PROCESSING ( WIRE - MANUAL - AND BACS) TO INCLUDE THE ONES ABOVE (see what i did there PL Re ) 
##~=====================================================================================================================================================
            $InterestignInvoices = Get-PnPListItem -Connection $Connection -List "Invoices" | Where-Object { $_.FieldValues._PaymentBatch -gt "" -and $_.FieldValues.wfSubStage -lt "9.0"  } | Sort-Object { $_.FieldValues._PaymentBatch } #that are at the right stage
            # do this because we may havbe changed the collection we searched before , ie moved some to stage 9 and renamed others and even backs matched them so lets refresh the workign set  
    
##~=====================================================================================================================================================
#~  BATCH STATUS ANALYSIS  _ AGAINST THE STAGES IM INTERESTED IN
##~=====================================================================================================================================================
            # firstly lets Get Some batch data insight into how they are doing against the specific stages may not be super elegant but its been very robust

            $Batch8_0 = @()
            $Batch8_1 = @()
            $Batch8_2 = @()
            $Batch8_3 = @()
            $thisBatch = ""
            foreach ($invoice in $InterestignInvoices) { # this is really a PROXY  for Each Batch in interestign invoices
                if ($invoice.FieldValues._PaymentBatch -ne $thisBatch) {
                    # lets address the status of each batch in turn this only works cos the invices are in batch order
                    $thisBatch = $invoice.FieldValues._PaymentBatch #well it does now!
            
                    $Batch8_0 += $(BatchStatus -Invoices $InterestignInvoices -Bt $thisBatch -St "8.0 Batched")
                    $Batch8_1 += $(BatchStatus -Invoices $InterestignInvoices -Bt $thisBatch -St "8.1 Matched")
                    $Batch8_2 += $(BatchStatus -Invoices $InterestignInvoices -Bt $thisBatch -St "8.2 First release")
                    $Batch8_3 += $(BatchStatus -Invoices $InterestignInvoices -Bt $thisBatch -St "8.3 Awaiting second release")
                    logActivity -Indent 0 -Type "2 Info" -Message  "Info1 | BACS Job |  |  | Evaluated $thisBatch  against statuses 8.0 to 8.3 "
                }
            }

            #log out the batches for DEBUG purposes 
            writeoutbatch  $Batch8_0 
            writeoutbatch  $Batch8_1 
            writeoutbatch  $Batch8_2 
            writeoutbatch  $Batch8_3 



##~=====================================================================================================================================================
#~  MANUAL AT stage 8.0  - SEND THE EMAIL  - no need to wait as the screen builds the batch as needed but i should let it be over an hour old age before i send
##~=====================================================================================================================================================
            foreach ($batch in $Batch8_0) {
                if ($batch.BatchType -eq "Manual" -or $batch.BatchType -eq "Wire" -or $batch.BatchType -eq "Wire Transfer") {
                    # scenarion ONE  all is good
                    if ($batch.Count -eq $batch.AtStage) { # TURNS OUT ITS IMPORTANT ALSO MAY USE VALUE > 0  
                        if ($batch.SendEmail -eq $true) {
                            logActivity -Indent 0 -Type "4 Action" -Message  "Action | BACS Job 8.0 |  |  |  $($batch.BatchName) : $($batch.BatchType),  No matching needed ($($batch.Count) items) - sending batch email to $($batch.Primary)"
                            $MailItem   = hydrateEmail -BatchStage "8.0_Manual" -BatchDetail $batch -AccountsEmail $accountsEmail -sitepath $path
                            $MailResult = smtp-SendEmail -RecipientEmails $batch.Primary -MessageText $MailItem.Message -Subject $MailItem.Subject -CC $accountsEmail -From $accountsEmail -BCC "tim.ellidge@pacificlifere.com" -EmailDebug $instance.EmailDebug 
                            BatchSetEmailFlag -BatchName $batch.BatchName -SendMail $false
                        } else {
                            logActivity -Indent 0 -Type "2 Info" -Message  "$($batch.BatchName) No Email needed " 
                        }
                    }
                }
            }

##=====================================================================================================================================================
# MATCHED BATCH STATUS WORK  - 8.1 SEND EMAIL IF COMPLETE ---  ONLY BACS BATCHES GET TO 8.1
#======================================================================================================================================================
            foreach ($batch in $Batch8_1) {
                # scenarion Failures
                if ($batch.Failed -gt 0) {
                    logActivity -Indent 0 -Type "2 Info" -Message  "BACS $($batch.BatchName) has $($batch.Failed) match failures and the last update was at $($batch.Latest) waiting..."
                } else {
                    if ($batch.NotProcessed -gt 0) {
                        logActivity -Indent 0 -Type "2 Info" -Message  "BACS $($batch.BatchName) has $($batch.NotProcessed) items not processed and the last update was at $($batch.Latest) waiting..."
                    } else {
                        if ($batch.Count -eq $batch.AtStage) {
                            if ($batch.SendEmail -eq $true) {
                                logActivity -Indent 0 -Type "4 Action" -Message  "Action | BACS Job 8.1 |  |  |  $($batch.BatchName) All matched OK ($($batch.Count) items) - sending 8.1 email to $($batch.Primary)"
                                $MailItem   = hydrateEmail -BatchStage "8.1_BACS_Matched" -BatchDetail $batch -AccountsEmail $accountsEmail -sitepath $path
                                $MailResult = smtp-SendEmail -RecipientEmails $batch.Primary -MessageText $MailItem.Message -Subject $MailItem.Subject -CC $accountsEmail -From $accountsEmail -BCC "tim.ellidge@pacificlifere.com" -EmailDebug $instance.EmailDebug
                                BatchSetEmailFlag -BatchName $batch.BatchName -SendMail $false
                            } else {
                                logActivity -Indent 0 -Type "2 Info" -Message  "$($batch.BatchName) No Email needed " 
                            }
                        }
                    } 
                }
            }

##=====================================================================================================================================================
#  PRE RELEASE BATCH STATUS ANALYSIS  _ 8.2 FOR BOTH BACS AND MANUAL AND WIRE (AWAITING FIRST RELEASE)
##=====================================================================================================================================================
            foreach ($batch in $Batch8_2) {
                if ($batch.Count -eq $batch.AtStage) {  
                    # TEST HERE To Go TO 8.3 OR 9.0
                    if ($batch.Secondary -gt "") {
                        # this little piggie has another hoop to jump through - so we will move it on and it will get picked up next time through 
                        logActivity -Indent 0 -Type "4 Action" -Message  "Action | BACS Job |  |  |  $($batch.BatchName) [$($batch.BatchType)] all pre released - Sending to secondary releaser"
                        $batchItems = $InterestignInvoices | Where-Object { $_.FieldValues._PaymentBatch -eq $batch.BatchName } # so pick up the associuated invoices
                        UpdateItemsTo -Items  $batchItems -toStage "8.3 Awaiting second release" -Action '<i class="far fa-clock fa-2x fa-fw iBlue"></i> Awaiting second Release' -Comment "Batch $($batch.BatchName) was pre released by $($batch.Primary)  and will be sent to  $($batch.Secondary)"  -EmailAlert $true
                    } else {
                        logActivity -Indent 0 -Type "4 Action" -Message  "Action | BACS Job |  |  |  $($batch.BatchName) [$($batch.BatchType)] complete - Setting all to released"
                        $batchItems = $InterestignInvoices | Where-Object { $_.FieldValues._PaymentBatch -eq $batch.BatchName } # so pick up the associuated invoices
                        UpdateItemsTo -Items  $batchItems -toStage "9.0 Released" -Action '<i class="fas fa-pound-sign fa-2x fa-fw iGreen"></i> Released' -Comment "Batch $($batch.BatchName) was released in full." -EmailAlert $false
                    }
                } else {
                    #<#
                    # so in this scenarion soem have any gone back to stage 1 (the rejected ones) but the processign on the batch is complete
                    if ($batch.Rejected -gt 0 ) {
                        $Rejects = ListInvoiceItems -List $InvoiceList -St "1.0 New Invoice" -Bt $batch.BatchName
                        # so i need to remove them from the batch and set the email to "no" on these items
                        $batchFailedItems = $InterestignInvoices | Where-Object { $_.FieldValues._PaymentBatch -eq $batch.BatchName -and $_.FieldValues.wfSubStage -eq "1.0 New Invoice" }
                        if ($batchFailedItems.length -gt 0) {
                            foreach ($RottenFailure in $batchFailedItems) {
                                # again maybe set the JSON to reflect this ?
                                $RottenFailure["_PaymentBatch"] = $null
                                $RottenFailure["AssignedTo1"] = $null
                                $RottenFailure["BusinessOwner"] = $null
                                $RottenFailure["_wfLockTime"] = $null
                                $RottenFailure["EmailAlert"] = $false
                                $RottenFailure.SystemUpdate()
                            }
                            #send an email to AP 
                            logActivity -Indent 0 -Type "5 Warning" -Message  "Action | BACS Job 8.2 |  |  |  Sending an email concerning $($batch.BatchName) these invoices are rejected $Rejects"
                            $MailItem   = hydrateEmail -BatchStage "8.2_BACS_Reject" -BatchDetail $batch -AccountsEmail $accountsEmail -sitepath $path
                            $MailResult = smtp-SendEmail -RecipientEmails $accountsEmail -MessageText $MailItem.Message -Subject $MailItem.Subject -CC $accountsEmail -From $accountsEmail -BCC "tim.ellidge@pacificlifere.com" -EmailDebug $instance.EmailDebug
                            $batchItems = $InterestignInvoices | Where-Object { $_.FieldValues._PaymentBatch -eq $batch.BatchName -and $_.FieldValues.wfSubStage -eq "8.2 First release" } # so pick up the associuated invoices
                            #so different in australia

                           if($instance.BU -eq "AU"){
                                if ($batch.Secondary -gt ""){
                                   UpdateItemsTo -Items  $batchItems -toStage "8.3 Awaiting second release" -Action '<i class="far fa-clock fa-2x fa-fw iBlue"></i> Awaiting second Release' -Comment "Batch $($batch.BatchName) was pre released by $($batch.Primary)  and will be sent to  $($batch.Secondary)"  -EmailAlert $true
                                } else {
                                   UpdateItemsTo -Items  $batchItems -toStage "9.0 Released" -Action '<i class="fas fa-pound-sign fa-2x fa-fw iGreen"></i>Partial  Released' -Comment "Batch $($batch.BatchName) was partially released." -EmailAlert $false
                                }
                           } else {     
                                UpdateItemsTo -Items  $batchItems -toStage "8.1 Matched" -Comment $null -EmailAlert $true 
                           }
                        }
                    } else {
                        logActivity -Indent 0 -Type "2 Info" -Message  "Info | BACS Job 8.2|  |  |  $($batch.BatchName) Is not matched or fully processed and none rejected"
                    }
                }
            }

##=====================================================================================================================================================
# MATCHED BATCH STATUS WORK  - 8.3 SEND EMAIL FOR SECONDARY 
##=====================================================================================================================================================
            foreach ($batch in $Batch8_3) {
                # only applicable for BACS batches 
                # scenarion ONE  all is good because it only gets to 8.3 by code and if its good then some testing is not needed  but lets be super sure 
           
                if ($batch.Under -eq $batch.Count) {
                    logActivity -Indent 0 -Type "2 Info" -Message  "Info | BACS Job 8.3 |  |  |  $($batch.BatchName) all items are under so it's not apropriate for 8.3 consideration" 
                } else {
                    if ($batch.Count -eq $batch.AtStage) {
                        if ($batch.SendEmail -eq $true) {
                            logActivity -Indent 0 -Type "2 Info" -Message  "Action | BACS Job 8.3 |  |  |  $($batch.BatchName) All matched OK ($($batch.Count) items) - sending final 8.3 email to $($batch.Secondary)"
                            $MailItem   = hydrateEmail -BatchStage "8.3_Secondary" -BatchDetail $batch -AccountsEmail $accountsEmail -sitepath $path
                            $MailResult = smtp-SendEmail -RecipientEmails $batch.Secondary -MessageText $MailItem.Message -Subject $MailItem.Subject -CC $accountsEmail -From $accountsEmail -BCC "tim.ellidge@pacificlifere.com" -EmailDebug $instance.EmailDebug
                            BatchSetEmailFlag -BatchName $batch.BatchName -SendMail $false
                        } else {
                            logActivity -Indent 0 -Type "2 Info" -Message  "Info | BACS Job 8.3 |  |  |  $($batch.BatchName) all at stage but No Email needed " 
                        }
                    } else {
                        # so Australia have a different rule set 
                        $limit = if($instance.BU -eq "AU"){$batch.Count} else {1} 
                        if ($batch.Rejected -gt 0) {
                            $Rejects = ListInvoiceItems -List $InvoiceList -St "8.2 First release" -Bt $batch.BatchName 
                            logActivity -Indent 0 -Type "2 Info" -Message  "$($batch.BatchName) has at least one item Rejected so moving batch to 8.1" 
                            $batchItems = $InterestignInvoices | Where-Object {$_.FieldValues._PaymentBatch -eq $batch.BatchName } # so pick up the associuated invoices
                            # update them and dont bother with the email as we are going to send a specific one here. 
                            # actually may be nice to pull out the comments if they exist on the rejected ones
                            UpdateItemsTo -Items  $batchItems -toStage "8.1 Matched" -Comment $null -EmailAlert $false  

                            $MailItem   = hydrateEmail -BatchStage "8.3_Secondary_Reject" -BatchDetail $batch -AccountsEmail $accountsEmail -sitepath $path
                            $MailResult = smtp-SendEmail -RecipientEmails $accountsEmail -MessageText $MailItem.Message -Subject $MailItem.Subject -From $accountsEmail -BCC "tim.ellidge@pacificlifere.com" -EmailDebug $instance.EmailDebug
                        } else {
                            If ($batch.Over -ge $limit) {
                                logActivity -Indent 0 -Type "2 Info" -Message  "Action | BACS Job |  |  |  $($batch.BatchName) [$($batch.BatchType)] complete - Setting all to released"
                                $batchItems = $InterestignInvoices | Where-Object { $_.FieldValues._PaymentBatch -eq $batch.BatchName } # so pick up the associuated invoices
                                UpdateItemsTo -Items  $batchItems -toStage "9.0 Released" -Action '<i class="fas fa-pound-sign fa-2x fa-fw iGreen"></i> Released' -Comment "Batch $($batch.BatchName) was released in full." 
                            } else {
                                logActivity -Indent 0 -Type "2 Info" -Message  "Info | BACS Job 8.3 |  |  |  $($batch.BatchName) not reached the threshold yet (and none rejected)" 
                            }
                        }  
                    }
                }
            }
        }
    }

#} catch{
#    write-host "OOPS" -ForegroundColor Magenta    logActivity -Indent 0 -Type "6 Error" -Message "UNABLE TO CONNECT TO SHAREPOINT" 
#}
#>


#======================================================================================================================================
#
# TASK END HERE - CLOSE OUT THE LOG AND REGISTER THE PING
#
#======================================================================================================================================
$p = basicPing

#WriteQuiet - only write a record IF the max error exceeds the limit   
logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished" -logAction "Write" # thsi does a log write in each loop so not needed here 

