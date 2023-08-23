Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "") 

function add-emailEvent() {
    #special case for AP notifications
    Param( [parameter(position = 1)] $InvoiceID, [parameter(position = 2)] $Recipient) #who is it goign to 

    $DeadLing  = Get-PnPListItem -Connection $sourceConnection -List "Invoices" -Identity $InvoiceID 

    logActivity -Indent 1 -Type "2 Info" -Message "SEND AN EMAIL to $Recipient" -ForegroundColor Cyan
    #we send an email by popping an entry into the Notifications list
    $NotificationValues = @{
        "Title"           = $DeadLing.FieldValues.Title;
        "MessageBody"     = "Not Used";
        "Recipient"       = $Recipient;
        "RecipientEmails" = $Recipient;
        "MessageLine"     = "Line Not used";
        "DAPSBU"          = $S.BU;
        'wfSubStage'      = $DeadLing.FieldValues.wfSubStage;
        "BusRef"          = $DeadLing.FieldValues.RefNo;
        "_wfFormType"     = "PLReInvoiceC";
        "From"            = $DeadLing.FieldValues.Editor.Email;
        "DAPSInvoiceNo"   = $DeadLing.FieldValues.RefNo;
    } 
    $n = Add-PnPListItem -Connection $sourceConnection -List "Lists/InvoiceNotifications" -Values $NotificationValues
    
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

# Where is the site and the library to store the LOG into 
$LogSiteURL = "https://pacificlife.sharepoint.com/sites/PLRe"
$LogLibName = "wfHistoryEvents"

# Who are we savign the log as and conenct to the log site as them 
$logaccountName = "svc_sp_sync@Pacificlifere.com" 
$logencrypted = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$logcredential = New-Object System.Management.Automation.PsCredential($logaccountName, $logencrypted)
$LogConnection = Connect-PnPOnline -Url $LogSiteURL -Credentials $logcredential -ReturnConnection

# Set up the Logging control data (static)
$Script:LogControl.JobName = $JobName;
$Script:LogControl.JobID   = $ID;
$Script:LogControl.LogLevel = "4 Action";
$Script:LogControl.LogLib = $LogLibName;
$Script:LogControl.LogConnection = $LogConnection;
$Script:LogControl.LogContact = "tim.ellidge@Pacificlifere.com";

# TYPES ARE : "1 Success", "2 Info", "3 Info", "4 Action", "5 Warning", "6 Error"
#Start the log with a simple entry
logActivity -Indent 0 -Type "1 Success" -Message "new root = $rt"
#======================================================================================================================================
#
# LOG SETUP END
#
#======================================================================================================================================
#

# Who are we connectign as to do the work? 
$accountName = "svc_sp_sync@Pacificlifere.com" 
$encrypted = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

#what sites are we going to return invoices too? 
$Sites = @{
    "UMe" = "https://pacificlife.sharepoint.com/sites/PLRe-UMeDAPS";
    "DC"  = "https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS";
    "AU"  = "https://pacificlife.sharepoint.com/sites/PLRe-AUDAPS";
}

#Lets reach out to the source site and get the connection 
$sourceConnection  = Connect-PnPOnline -Url "https://pacificlife.sharepoint.com/sites/PLRe-tDAPSApprovals" -Credentials $credential -ReturnConnection 
$SPInvoiceUpdates  = Get-PnPListItem -Connection $sourceConnection  -List "InvoiceLog" #-ErrorAction SilentlyContinue -ErrorVariable ErrVar
logActivity -Indent 0 -Type "1 Success" -Message  "found  $($SPInvoiceUpdates.Count) Invoices "

#Lets get the items from the Queue sorted by business unit and Invoice No ie DC|345 just so we do all the ones together and don't have to keep reconnecting 
$SortedUpdates = $SPInvoiceUpdates | Sort-Object -Property {$_.FieldValues.RefNo} 

logActivity -Indent 0 -Type "2 Info" -Message  "found  $($SortedUpdates.Count) Items to process"
$DestConnection = $null

if($SortedUpdates.Count -gt 0){
    foreach ($SU in $SortedUpdates) {
        #------------------------------------------------------------------------------------------------
        # GET SOME REFERENCE DATA
        #------------------------------------------------------------------------------------------------
        $BU      = $SU.FieldValues.RefNo.split("|")[0]
        $DAPSID  = $SU.FieldValues.RefNo.split("|")[1]
        $xAction = if($SU.FieldValues.Action) {[String] $SU.FieldValues.Action[0]} else {"0"}
        $xStage  = [String] $SU.FieldValues.wfSubStage[0];

        If ($DestConnection.Url -ne $Sites[$BU]){
            #Connect if we need to
            $DestConnection = Connect-PnPOnline -Url $Sites[$BU] -Credentials $credential -ReturnConnection 
            $ServerRelativeRoot = $Sites[$BU].replace("https://pacificlife.sharepoint.com","")
        }

        logActivity -Indent 0 -Type "2 Info" -Message  "Processing $($SU.FieldValues.RefNo) - $($SU.FieldValues.wfSubStage)  $($SU.FieldValues.Action) "   -ForegroundColor Cyan

        #------------------------------------------------------------------------------------------------
        #  so lets figure out the Action and progress / Delete / save as needed ALSO USer Field 4 
        #------------------------------------------------------------------------------------------------
        If($xAction -eq "2"){
            #Return to AP
            $nextStage = "1.0 New Invoice" # we will use this later
            $Values = @{ 
                "wfSubStage"          = $nextStage
                "_wfStatusChangeDate" = $SU.FieldValues._wfTime
                "_wfLockTime"         = $null
                "_UserField4"         = $SU.FieldValues._UserField4
            }
            $u2 = set-PnPListItem -Connection $DestConnection -List "Invoices" -Identity $DAPSID -Values $Values
            
            #DELETE THE ITEM AT THIS STAGE FROM 0365 (don't delete the apportionments)
            $d = Remove-PnPListItem -Connection $sourceConnection -List "Invoices" -Identity $SU.FieldValues.LocalInvoiceID -Recycle -Force
            $actionLog = '<i class="far fa-id-badge fa-fw fa-2x iRed"></i> Returned to AP'
            logActivity -Indent 0 -Type "4 Action" -Message "Removed Invoice $($SU.FieldValues.RefNo) its back with AP - send an email "
            #
            # TODO SEND AN AP EMAIL  AND ADD TO WF HISTORY ??? or is that part of the process later on LETS SEE 
            #
        } else {
            if ($xAction -eq "1"){
                #Progress to the next stage :-) BUT Which One ? 
                if($xStage -eq "2"){ 
                    $nextStage   = "3.0 Waiting authorisation"
                    $actionLog   = '<i class="far fa-id-badge fa-fw fa-2x iAmber"></i> Endorsed'
                    $StageAction = "Authorisation"
                    $assignedto  = $SU.FieldValues.AuthoriserEmail
                } else {
                    if($xStage -eq "3"){
                        if($SU.FieldValues.SecAuthoriserEmail -gt "" ){
                            $nextStage   = "4.0 Waiting secondary authorisation"
                            $actionLog   = '<i class="far fa-check-square fa-fw fa-2x iGreen"></i> Authorised (primary)'
                            $StageAction = "Secondary Authorisation"
                            $assignedto  = $SU.FieldValues.SecAuthoriserEmail
                        } else {
                            $nextStage   = "5.0 Coding in PeopleSoft"
                            $actionLog   = '<i class="far fa-check-square fa-fw fa-2x iGreen"></i> Authorised'
                            $StageAction = "" 
                            $assignedto  = "" 
                        }
                    } else {
                        $nextStage   = "5.0 Coding in PeopleSoft"
                        $actionLog   = '<i class="far fa-check-square fa-fw fa-2x iGreen"></i> Authorised (secondary)'
                        $StageAction = "" 
                        $assignedto  = "" 
                    }
                }
                ##I think it goes here to update the source record 
                #update the item in O365 in dest and souce (no need to wait for a cycle)

                #Need to get the current wfstatus changedate
                $I  = Get-PnPListItem -Connection $sourceConnection -List "Invoices" -Id $SU.FieldValues.LocalInvoiceID
                $ts = New-TimeSpan -Start $I.FieldValues._wfStatusChangeDate -End $SU.FieldValues._wfTime
                #write-host $ts.TotalHours -ForegroundColor Red

                $O365Values = @{ 
                    "wfSubStage"          = $nextStage;
                    "_wfStatusChangeDate" = $SU.FieldValues._wfTime;
                    "DAPSUsed"            = $null;
                    "wfSubStage1"         = $($StageAction);
                    "AssignedTo1"         = $assignedto;
                }
                $u  = set-PnPListItem -Connection $sourceConnection -List "Invoices" -Identity $SU.FieldValues.LocalInvoiceID -Values $O365Values

            } else { #xAction must be null or zero 
                # we are just goign to add a comment but leave at the same stage (dotn bother to update the record 
                $nextStage= $SU.FieldValues.wfSubStage
                $actionLog = '<i class="far fa-save fa-fw fa-2x iBlue"></i> Saved'
            }

            #update the destination AP Site with the data needed don't clear the lock date else it will come back to us again 
            $SP2013Values = @{ 
               "wfSubStage"          = $nextStage
               "_wfStatusChangeDate" = $SU.FieldValues._wfTime
               "_UserField4"         = $SU.FieldValues._UserField4
            }
            $u2 = set-PnPListItem -Connection $DestConnection   -List "Invoices" -Identity $DAPSID -Values $SP2013Values
        }
        logActivity -Indent 0 -Type "2 Info" -Message  "this stage no $($xStage) Action $($actionLog) psComment:$($SU.FieldValues._UserField4)" -ForegroundColor Green

        #------------------------------------------------------------------------------------------------
        # Manage the Apportionments
        #------------------------------------------------------------------------------------------------
        
        $ARN = $SU.FieldValues.RefNo
        if($ARN.length -gt 5){
            $SourceItemGUIDS = @();
                
            logActivity -Indent 0 -Type "2 Info" -Message  "Querying Apportionments" 
            #get the Apportionments from DAPS Hybrid
            $APPS = Get-PnPListItem -Connection $sourceConnection -List "Lists/InvoiceApportionments" -pagesize 2000 | Where-Object { $_.FieldValues.BusRef -eq $ARN }  #  -ErrorAction SilentlyContinue -ErrorVariable ErrVar       
    
            if($APPS){
                logActivity -Indent 0 -Type "2 Info" -Message  "$($APPS.count) Apportionment records" -ForegroundColor Yellow
                foreach($AP in $APPS){
                    #BUILD AN OBJECT FOR THIS to hold its METADATA

                    $thisApportionment = @{
                        "Title"           = $($AP.FieldValues.Title); 
                        "BusRef"          = $($AP.FieldValues.BusRef);
                        "APPCategory"     = $($AP.FieldValues.APPCategory);
                        "APPBU"           = $($AP.FieldValues.APPBU);
                        "APPMin"          = $($AP.FieldValues.APPMin);
                        "APPMax"          = $($AP.FieldValues.APPMax);
                        "APPCode"         = $($AP.FieldValues.APPCode);
                        "APPAmount"       = $($AP.FieldValues.APPAmount);
                        "_wfUser"         = $($AP.FieldValues._wfUser.Email);
                        "_wfTime"         = $($AP.FieldValues._wfTime);
                        "APPDescription"  = $($AP.FieldValues.APPDescription);
                        "_SourceItemGUID" = $AP.FieldValues.GUID;
                    }
                    $SourceItemGUIDS += $AP.FieldValues.GUID;
                
                    $folderId  = $ARN.Replace("|","-")         
                    #Do a search to get the matching Source site Apportionments look in the FOLDER AND CHECK BY GUID
                    $DestAPP = Get-PnPListItem -Connection $DestConnection  -List "Lists/InvoiceApportionments"-PageSize 1000 -FolderServerRelativeUrl "$($ServerRelativeRoot)/Lists/InvoiceApportionments/$($folderId)" | Where-Object { $_.FieldValues._SourceItemGUID -eq $AP.FieldValues.GUID }
                    # did we find one ? 
                    if($DestAPP){
                        # to do a test of the _wfTime should decide if we can leave it alone or update it 
                        if ($DestApp.FieldValues._wfTime -eq $AP.FieldValues._wfTime){
                            logActivity -Indent 0 -Type "2 Info" -Message  "No Need  to update $($AP.FieldValues.Title)"
                        } else {
                            logActivity -Indent 0 -Type "2 Info" -Message  "Updating Apportionment $($AP.FieldValues.Title)"
                            $a = Set-PnPListItem -Connection $DestConnection -List "Lists/InvoiceApportionments" -Identity $DestAPP.Id -Values $thisApportionment
                        }
                    } else {
                        logActivity -Indent 0 -Type "2 Info" -Message  "Adding Apportionment $($AP.FieldValues.Title)"
                        #THIS CREATES A FOLDER IF NEEDED 
                        $a = Add-PnPListItem -Connection $DestConnection -List "Lists/InvoiceApportionments" -Values $thisApportionment -Folder $folderId
                    }   
                }    
            }

            #Do the delete processing not so cute but there are so few of them for any given invoice so we can be clunky and not clever
            #TEST FOLDEER ID
            if ($folderId -ne ""){
                $DestAPPS = Get-PnPListItem -Connection $DestConnection -List "Lists/InvoiceApportionments" -FolderServerRelativeUrl "$($ServerRelativeRoot)/Lists/InvoiceApportionments/$($folderId)"
                if($DestAPPS){
                    ForEach($DA in $DestAPPS){
                        $rowID = $DA.Id
                        if($rowID -is [Int] -and $rowID -gt 0){
                            if ($DA.FieldValues._SourceItemGUID -notin $SourceItemGUIDS){
                                logActivity -Indent 0 -Type "2 Info" -Message  "Should delete $($rowID) -   $($DA.FieldValues.Title)"
                                # $d = Remove-PnPListItem -Connection $DestConnection -List "Lists/InvoiceApportionments" -Identity $rowID -Recycle -Force
                            }
                        } else {
                            logActivity -Indent 0 -Type "5 Warning" -Message  "$($rowID) -   $($DA.FieldValues.Title) had a NULL ID? "
                        }
                    }
                } else {
                    logActivity -Indent 0 -Type "5 Warning" -Message  "/InvoiceApportionments/$($folderId) - has no items in it"
                }
            } else {
                logActivity -Indent 0 -Type "6 Error " -Message  "/InvoiceApportionments/$($folderId) - IS FUCKED"
            }
        } else {
            logActivity -Indent 0 -Type "6 Error" -Message  "$($SU.FieldValues.Title) - has no Bus Ref  in it"
        }

        #------------------------------------------------------------------------------------------------
        # Add a WF History Item
        #------------------------------------------------------------------------------------------------ 
        $HistoryRecord = @{
            "Title"	          = "$($DAPSID) - $($SU.FieldValues.wfSubStage)";
            "_wfAction"	      = $actionLog;
            "_wfFormID"	      = $DAPSID;
            "_wfFormType"	  = "PLREInvoiceC";
            "_wfLogComment"	  = $SU.FieldValues.Comment;
            "_wfLongComment"  = $SU.FieldValues.Comment;
            "_wfPrevStage"	  = $SU.FieldValues.wfSubStage;
            "_wfStageChange"  = $nextStage;	
            "_wfStreamStatus" = "$($nextStage)||";	
            "_wfStreamTime0"  = $ts.TotalHours;	
            "_wfTime"         = $SU.FieldValues._wfTime;	
            "_wfUser"	      = $SU.FieldValues._wfUser.LookupValue;
            "UserLogData"     = "";
        }
        ## the folder needs to be F[Id]for GO Live
        logActivity -Indent 0 -Type "4 Action" -Message  "Adding a history Record for $($DAPSID) - $($SU.FieldValues.wfSubStage)"
        $newHist = Add-PnPListItem -Connection $destConnection -List "Lists/wfHistoryF" -Folder "F$($DAPSID)" -Values $HistoryRecord

        #------------------------------------------------------------------------------------------------
        # Finally Remove the INVOICE LOG ITEM
        #------------------------------------------------------------------------------------------------
        $u = set-PnPListItem    -Connection $sourceConnection -List "InvoiceLog" -Identity $SU.Id -Values @{"Title" = "$($SU.FieldValues.RefNo) $($SU.FieldValues.wfSubStage)"} 
        logActivity -Indent 0 -Type "4 Action" -Message  "Removing the log item to the recycle bin"
        $d = Remove-PnPListItem -Connection $sourceConnection -List "InvoiceLog" -Identity $SU.Id -Recycle -Force
    }
} else {
    logActivity -Indent 1 -Type "2 Info" -Message "No Invoices to process - Nothing in the invoice log"
} 


#======================================================================================================================================
#
# TASK END HERE - CLOSE OUT THE LOG AND REGISTER THE PING
#
#======================================================================================================================================
$JobDuration = ((get-date) - $Script:LogControl.LogFirstCall).TotalSeconds # how long did the processing take
$ping = @{"LastAlive" = get-date ; "Duration(s)" = $JobDuration ; }

#check if its the first one if so create a directory otherwise just save it  
if ((test-path -Path ".\9-PINGS\$($JobName)") -eq $false) { New-Item -Path ".\9-PINGS\$($JobName)" -ItemType directory }
$a = $ping | Out-File -FilePath ".\9-PINGS\$($JobName)\$($JobName)$(Get-date -Format "yyMMdd-HH_mm").txt" 

#WriteQuiet - only write a record IF the max error exceeds the limit $Script:LogControl.LogLevel 
#Write always writes at lease one line  
logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished in $($JobDuration) seconds" -logAction "WriteQuiet"
#>


