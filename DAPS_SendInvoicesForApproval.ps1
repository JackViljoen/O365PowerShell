Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "") 



#############################################################################################################
function get-wfHistoryAsHTML() {
    Param( [parameter(position = 1)] $INVID, [parameter(position = 2)] $SitePath , [parameter(position = 3)] $Cnx) #the Invoice i want to get the History for
    # ============================================================
    # USE the REST method - turns out its sub second response as opposed to 20 or 30 seconds and loads of memory used up 
    # note the use of the backtick nefore the $filter=  its doesn use the filter if yoy leave it out  

    $output = Invoke-PnPSPRestMethod -Connection $Cnx  -Url "$($SitePath)/_api/Web/Lists/getByTitle('_wfHistoryF')/items?`$filter=%20OData__wfFormID%20eq%20$($INVID)"
    $SPHistory = $output.value #allocate it to a variable 
    $HistoryTable = ""
    $attachments = @()

    #note the object uses the full OData_ names for the fields as its not the usual return type it doesn't have ".FieldValues"
    if ($SPHistory) {
        logActivity -Indent 1 -Type "2 Info" -Message "Found $($SPHistory.count) wf History records" -ForegroundColor Yellow
        foreach ($SPH in $SPHistory) {
            $Icon = ""
            if($SPH.Attachments) {
                $Icon = "<div class='attachmment' data-item='$($SPH.ID)' style='float:right;'> XXX$($SPH.ID)XXX </div>"
                $attachments+= $SPH.ID
            }
            if ($SPH.OData__wfAction.indexOf("far fa-save fa-fw fa-2x iBlue") -eq -1 -or $null -ne $SPH.OData__wfLogComment -or $SPH.Attachments) {
                $HistoryTable += "<div class='hitem'><div class='hitemheader'style=style='float:left; width:100%'>$(get-date($SPH.OData__wfTime) -format "dd MMM HH:mm") $($SPH.OData__wfUser) : $($SPH.OData__wfAction.replace('fa-2x',''))  $($Icon) </div><div class='hcomment'>$($SPH.OData__wfLogComment) </div></div>"
            }
        }
    }
    #pass back the HTML AND the IDS of any items with attachments
    return @{thisHTML = $HistoryTable; attach = $attachments}
}
###############################################################################################################

function CheckPermissions() {
    Param( [parameter(position = 1)] $SOURCE, [parameter(position = 2)] $DEST, [parameter(position = 3)] $MANAGERS) # the Invoices i want to compare Securty For
   
    $Action = ""
    #Build an arrays  of the people and owner group.. It cant have any empty slots  
    $SourcePeopleRoles  = @();
    if ($SOURCE.FieldValues._EndorserEmail)      {$SourcePeopleRoles += ($SOURCE.FieldValues._EndorserEmail -replace '<[^>]+>', '').ToLower() + "|Contribute"}
    if ($SOURCE.FieldValues._AuthoriserEmail )   {$SourcePeopleRoles += ($SOURCE.FieldValues._AuthoriserEmail -replace '<[^>]+>', '' ).ToLower() + "|Contribute"}
    if ($SOURCE.FieldValues._SecAuthoriserEmail) {$SourcePeopleRoles += ($SOURCE.FieldValues._SecAuthoriserEmail -replace '<[^>]+>', '').ToLower() + "|Contribute"}
    $SourcePeopleRoles += $MANAGERS.LoginName + "|Contribute"

    $DestPeopleRoles = @();
   
    $perms = Get-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $DEST.Id
   
    foreach ($P in $perms.Permissions){
        if($P.PrincipalType -eq "SharePointGroup"){
            foreach($PL in $P.PermissionLevels){
                $DestPeopleRoles += ($P.PrincipalName+"|"+$PL.Name) 
            }
        } else {
            if($P.PrincipalType -eq "User"){ 
                foreach($PL in $P.PermissionLevels){
                    #TODO MUST BE A BETTER WAY 
                    $DestPeopleRoles += ($P.PrincipalName.split("|")[2].split("#")[0].Replace("_","@")+"|"+$PL.Name) #fish out the email or the Group
                }
            }
        }
    }
        
    #this is my new go-to comparison pattern 
    $Diffs = Compare-Object -ReferenceObject $SourcePeopleRoles  -DifferenceObject $DestPeopleRoles
    if ($Diffs) {       
        foreach ($D in $Diffs) {
            $DBits = $D.InputObject.split("|") 

            if ($D.SideIndicator -eq "<=") {
                #do the adds
                if($DBits[0].IndexOf("@") -eq -1){
                    #its a group
                    $a = Set-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $DEST.Id -AddRole $DBits[1] -Group $DBits[0] -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                } else {
                    # I HATE THIS CODE !!! ITS HORRID 
                    if($DBits[0].indexOf("'") -gt -1) {
                        $n = $DBits[0].split("@")[0]
                        $n = "$($n.split(".")[1]), $($n.split(".")[0])"
                        $DBits[0] = $n
                    }
                    $a = Set-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $DEST.Id -AddRole $DBits[1] -User $DBits[0] -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                }
                if ($Errvar[0].Count -gt 0) {
                    logActivity -Indent 1 -Type "5 Warning" -Message "Problem Adding person $($D.InputObject)"
                } else {
                    $Action += "Added $($D.InputObject) | "
                }
            } else {
                #Do the removes dont forget we cant remove limited access  
                if($DBits[1] -ne "Limited Access"){
                    if($DBits[0].IndexOf("@") -eq -1){
                        #its a group
                        $a = Set-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $DEST.Id -RemoveRole $DBits[1] -Group $DBits[0] -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                    } else {
                        $a = Set-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $DEST.Id -RemoveRole $DBits[1] -User $DBits[0] -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                    }
                    if ($Errvar[0].Count -gt 0) {
                        logActivity -Indent 1 -Type "5 Warning" -Message "Problem removing role from $($D.InputObject)"
                    } else {
                        $Action += "removed $($D.InputObject) | "
                    }
                }
            }        
        }
    }
    if ($Action -eq "") { $Action = "No permission changes needed" }
    return $action
}

function CheckPermissionNames() {
    Param( [parameter(position = 1)] $SOURCE, [parameter(position = 2)] $DEST, [parameter(position = 3)] $MANAGERS) # the Invoices i want to compare Securty For
   
    $Action = ""
    #Build a couple of arrays (it may have some empty slots) will that mattter - turns out no it doesnt ?  
    $SourcePeople = @();
    $SourcePeople += ($SOURCE.FieldValues._EndorserEmail -replace '<[^>]+>', '').ToLower()
    $SourcePeople += ($SOURCE.FieldValues._AuthoriserEmail -replace '<[^>]+>', '' ).ToLower()
    $SourcePeople += ($SOURCE.FieldValues._SecAuthoriserEmail -replace '<[^>]+>', '').ToLower()

    $DestPeople = @();
    if($DEST.FieldValues._InvEndorser.Email)     {$DestPeople += ($DEST.FieldValues._InvEndorser.Email).ToLower()}
    if($DEST.FieldValues._InvAuthorise.Email)    {$DestPeople += ($DEST.FieldValues._InvAuthorise.Email).ToLower()}
    if($DEST.FieldValues._InvSecAuthorise.Email) {$DestPeople += ($DEST.FieldValues._InvSecAuthorise.Email).ToLower()}
        
    #this is my new go-to comparison pattern 
    $Diffs = Compare-Object -ReferenceObject $SourcePeople  -DifferenceObject $DestPeople
    if ($Diffs) {       
        foreach ($D in $Diffs) {
            if ($D.InputObject) {
                #it may be an empty string so skip it 
                if ($D.SideIndicator -eq "<=") {
                    #do the adds
                    $a = Set-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $O365File.Id -AddRole "Contribute" -User $D.InputObject -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                    if ($Errvar[0].Count -gt 0) {
                        logActivity -Indent 1 -Type "5 Warning" -Message "Problem Adding person $($D.InputObject)"
                    } else {
                        $Action += "added $($D.InputObject) | "
                    }
                }
                else {
                    #Do the removes 
                    $a = Set-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $O365File.Id -RemoveRole "Contribute" -User $D.InputObject -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                    if ($Errvar[0].Count -gt 0) {
                        logActivity -Indent 1 -Type "5 Warning" -Message "Problem removing person $($D.InputObject)"
                    } else {
                        $Action += "removed $($D.InputObject) | "
                    }
                }
            }
        }
    }

    $a = Set-PnPListItemPermission -Connection $destConnection -List "Invoices" -Identity $O365File.Id -AddRole "Contribute" -Group $MANAGERS -ErrorAction SilentlyContinue -ErrorVariable ErrVar
    if ($Errvar[0].Count -gt 0) {
        logActivity -Indent 1 -Type "5 Warning" -Message "Problem Adding group  $($MANAGER)"
    } else {
        $Action += "added $($MANAGERS) | "
    }


    if ($Action -eq "") { $Action = "No permission changes needed" }
    return $action
}

function CheckApportionments() {
    Param( [parameter(position = 1)] $InvoiceData) # the Invoices i want to check apportionments on 

    $ABU = $InvoiceData.Business
    $ARN = $InvoiceData.RefNo

    $Apps = Get-PnPListItem -Connection $destConnection -List "Lists/InvoiceApportionments" | Where-Object { $_.FieldValues.BusinessUnit -eq  $ABU -and $_.FieldValues.BusRef -eq $ARN -and $_.FieldValues.APPCategory -eq "Additional Cost Centre" }
    if($Apps.count -gt 0){
        logActivity -Indent 1 -Type "2 Info" -Message "Already have an Additional CC apportionment not adding a default"
    } else {
        if($InvoiceData._InvoiceDept){
            logActivity -Indent 1 -Type "2 Info" -Message  "$($InvoiceData.Title) : $($InvoiceData._InvoiceDept) : $($InvoiceData.RefNo) : $($InvoiceData.DAPSBU)"
            logActivity -Indent 1 -Type "2 Info" -Message "   MISSING CC RECORD les make one up " -ForegroundColor RED
            ## the departments are different in the Additional CC - I wish they werent but they are... maybe a fix for later
            $DeptName = ($InvoiceData._InvoiceDept).Split("|")[1].Trim() +" "+($InvoiceData._InvoiceDept).Split("|")[0].Trim()
            logActivity -Indent 1 -Type "2 Info" -Message $DeptName -ForegroundColor Cyan
            
            #Build the apportionment record to write into 0365
            $ApportionmentRecord = @{
                "Title"          = $InvoiceData.RefNo + " Defaut Value";
                "BusinessUnit"   = $InvoiceData.Business; #TOdo WTF !!! this isnt correct -- actually it may be ok 
                "APPCategory"    = "Additional Cost Centre";
                "APPCode"        = $DeptName
                "APPDescription" = $DeptName +"|"+ $InvoiceData._InvAuthorise
                "APPAmount"      = $InvoiceData.InvoiceAmount;
                "APPSequence"    = 0;
                "APPMin"         = 0;
                "APPMax"         = 10;
                "BusRef"         = $InvoiceData.RefNo;
                "_wfTime"        = $InvoiceData._wfStatusChangeDate;
                "_wfUser"        = $InvoiceData._InvAuthorise;
            }

            logActivity -Indent 1 -Type "2 Info" -Message "`t`t`tAdding apportionment record"
            $newAPP = Add-PnPListItem -Connection $destConnection -List "Lists/InvoiceApportionments" -Values $ApportionmentRecord

        } else {
            logActivity -Indent 1 -Type "2 Info" -Message "Record $($InvoiceData.Title) -  $($InvoiceData.RefNo) has no dept code !!"
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

# Who are we connectign as to do the work? 
$accountName = "svc_sp_sync@Pacificlifere.com" 
$encrypted   = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$credential  = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

#what sites are we going to scan? 
$Sites = @(
    @{"Active" = $true;  "SiteUrl"  = "https://pacificlife.sharepoint.com/sites/PLRe-UMEDAPS"; "BU" = "UME"; "From" = "accounts@underwriteme.co.uk" };
    @{"Active" = $true;  "SiteUrl"  = "https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS";  "BU" = "DC";  "From" = "accountspayableUK@pacificlifere.com" };
    @{"Active" = $true;  "SiteUrl"  = "https://pacificlife.sharepoint.com/sites/PLRe-AUDAPS";  "BU" = "AU";  "From" = "Accounts.PayableAUS@PacificLifere.com" };
)

#Lets reach out to the destination site and get the connection and the roles
$destConnection = Connect-PnPOnline -Url "https://pacificlife.sharepoint.com/sites/PLRe-tDAPSApprovals" -Credentials $credential -ReturnConnection 
#dest stuff for working with the security is this needed
$roles = Get-PnPRoleDefinition -Connection $DestConnection

forEach ($S in $sites) {
    If ($S.Active) {
        logActivity -Indent 1 -Type "2 Info" -Message " SCANNING =  $($S.SiteUrl)" -ForegroundColor Yellow 
        #get the relevent BU security group 
        $ManagerGroup = Get-PnPGroup -Connection $DestConnection -Identity "DAPS Hybrid $($S.BU) Managers"
         
        # CONNECT TO THE SOURCE SHAREPOINT
        $relativeRoot = $S.SiteUrl.Replace("https://pacificlife.sharepoint.com", "")
        $thisConnection = Connect-PnPOnline -Url $S.SiteUrl -Credentials $credential -ReturnConnection 
     
        logActivity -Indent 1 -Type "2 Info" -Message "Looking for invoices in $($thisConnection.Url) relative root $($relativeRoot)" -ForegroundColor Green
      
        $LockTime = Get-Date  
 
        # GET ALL UNLOCKED INVOICES AT STAGES 2 to 8
        $SPInvoices = Get-PnPListItem -Connection $thisConnection  -List "Invoices" | Where-Object { $_.FieldValues.wfSubStage -gt 2 -and $_.FieldValues.wfSubStage -lt 5 -and $_.FieldValues._wfLockTime -eq $null }
        logActivity -Indent 1 -Type "2 Info" -Message "found  $($SPInvoices.Count) Invoices"
        foreach ($SPInvoice in $SPInvoices) {
            if ($SPInvoice.FieldValues.File_x0020_Type -eq "pdf") {
                # IF ITS A PDF
                $xStage = $SPInvoice.FieldValues.wfSubStage.Substring(0, 1);

                #figure out who it is assigned too 
                switch ($xStage) {
                    2 { $assignedToEmail = $SPInvoice.FieldValues._EndorserEmail -replace '<[^>]+>', ''     ; $StageAction = "Endorsement" }
                    3 { $assignedToEmail = $SPInvoice.FieldValues._AuthoriserEmail -replace '<[^>]+>', ''   ; $StageAction = "Authorisation" }
                    4 { $assignedToEmail = $SPInvoice.FieldValues._SecAuthoriserEmail -replace '<[^>]+>', ''; $StageAction = "Secondary Authorisation" }
                    default { $assignedToEmail = "" ; $StageAction = "" }
                }

                logActivity -Indent 1 -Type "2 Info" -Message "Trying $($SPInvoice.FieldValues._UserField2) - ($xStage) $($SPInvoice.FieldValues.wfSubStage)  $($SPInvoice.FieldValues.FileRef)  Assigned to $($assignedToEmail)"   -ForegroundColor Cyan
             
                #prepare the place its going to 
                $Path = $SPInvoice.FieldValues.FileRef
                $FileName = $SPInvoice.FieldValues.FileLeafRef
                $DestFolderLocation = "/sites/PLRe-tDAPSApprovals/Invoices/$($S.BU)"
                $DestPath = $DestFolderLocation + "/" + $FileName

                #Is there a file already there ? 
                $O365File = Get-PnPFile -Connection $destConnection -Url $DestPath -AsListItem   -ErrorAction SilentlyContinue
                
                #Check if File already exists 
                If ($O365File) {
                    logActivity -Indent 1 -Type "2 Info" -Message " $($fileName) already exists in O365 " 
                }  else {
                    logActivity -Indent 1 -Type "2 Info" -Message "FILE NOT THERE NEEDS ADDING " 
                    # Send the file to O365
                    $File = Copy-PnPFile -SourceUrl $Path  -TargetUrl "/sites/PLRe-tDAPSApprovals/Invoices/$($S.BU)"  -Overwrite -Force   -Connection $thisConnection
                    # Get the file back to start working with it
                    $O365File = Get-PnPfile -Connection $destConnection -Url $DestPath -AsListItem
                    logActivity -Indent 1 -Type "4 Action" -Message "Copied file to /Invoices/$($S.BU)"
                }
            
                #So set the security by comparign the Metadata before we update it - will this work ?  
                $SecurityAction = CheckPermissions -SOURCE $SPInvoice -DEST $O365File -MANAGERS $ManagerGroup
                logActivity -Indent 1 -Type "2 Info" -Message "Permissions check: $($SecurityAction)"


################################################################################################
                #Pull out the history for the  invoice in the source site as a HTML string
                $HistoryHTMLObj = get-wfHistoryAsHTML -SitePath $S.SiteUrl -INVID $SPInvoice.Id -Cnx $thisConnection
                $HistoryHTML = $HistoryHTMLObj.thisHTML

                #DO WE HAVE ANY ATTACHMENTS? 
                if($HistoryHTMLObj.attach){
                    foreach($AT in $HistoryHTMLObj.attach){
                        # get the attachmwent prtepend the name and send over to the O365 location attachments/BU/Ann-filename 
                        write-host $AT
                        Get-PnPListItemAttachment -connection $thisConnection -List "Lists/WfHistoryF" -Identity $AT -Path "C:\O365PowerShell\1-APPLICATIONS\O365DAPS\Attachments" -Force
                        $AttachmentFiles =  get-ChildItem -Path "C:\O365PowerShell\1-APPLICATIONS\O365DAPS\Attachments" 
                        #there will only be one but it has to go in a loop 
                        foreach($AF in $AttachmentFiles){  
                            Add-PnPFile -Connection $destConnection -Path $AF.FullName -Folder "Attachments/$($S.BU)" -Values @{BusRef=$AT} -NewFileName "A$($AT)-$($AF.Name)"
                            start-sleep 2
                            remove-Item $AF.FullName
                        }
                        #now replace the holder with a HREF 
                        $HistoryHTML = $HistoryHTML.Replace("XXX$($AT)XXX", "<a href='/sites/PLRe-tDAPSApprovalsUAT/Attachments/$($S.BU)/A$($AT)-$($AF.Name)' target='_blank'> View attachment </a>")
                    }
                }
###################################################################################

                #Should we send an Email ? has the assigned to changed
                If ($assignedToEmail -ne "" -and $assignedToEmail -ne $O365File.FieldValues.AssignedTo1.Email ){
                    logActivity -Indent 1 -Type "2 Info" -Message "SEND AN EMAIL to $assignedToEmail" -ForegroundColor Cyan
                    #we send an email by popping an entry into the Notifications list
                    $NotificationValues = @{
                        "Title"           = $SPInvoice.FieldValues.Title;
                        "MessageBody"     = "Not Used";
                        "Recipient"       = $assignedToEmail;
                        "RecipientEmails" = $assignedToEmail;
                        "MessageLine"     = "Line Not used";
                        "DAPSBU"          = $S.BU;
                        'wfSubStage'      = $SPInvoice.FieldValues.wfSubStage;
                        "BusRef"          = "$($S.BU)|$($SPInvoice.Id)";
                        "_wfFormType"     = "PLReInvoiceC";
                        "From"            = $S.From;
                        "DAPSInvoiceNo"   = $($SPInvoice.FieldValues.RefNo);
                    } 
                    $n = Add-PnPListItem -Connection $destConnection -List "Lists/InvoiceNotifications" -Values $NotificationValues
                }

                #Build the data to add the metadata 
                $thisItem = @{
                    '_InvAuthorise'       = $SPInvoice.FieldValues._AuthoriserEmail -replace '<[^>]+>', '' ;
                    '_InvEndorser'        = $SPInvoice.FieldValues._EndorserEmail -replace '<[^>]+>', '';
                    '_InvoiceDept'        = $SPInvoice.FieldValues.InvDept;
                    '_InvSecAuthorise'    = $SPInvoice.FieldValues._SecAuthoriserEmail -replace '<[^>]+>', '' ;
                    '_PayeeAccountNo'     = $SPInvoice.FieldValues._PayeeAccountNo;
                    '_PayeeSortCode'      = $SPInvoice.FieldValues._PayeeSortCode;
                    '_PaymentType'        = $SPInvoice.FieldValues._PaymentType;
                    '_SourceItemGUID'     = $SPInvoice.FieldValues.GUID;
                    '_SplitAmount'        = $SPInvoice.FieldValues.InvoiceAmount;
                    '_SystemUpdateDate'   = $SPInvoice.FieldValues._SystemUpdateDate;
                    '_UserField1'         = $SPInvoice.FieldValues._UserField1;
                    '_UserField4'         = $SPInvoice.FieldValues._UserField4;
                    '_Vendor'             = $SPInvoice.FieldValues._Vendor;
                    '_wfHistory'          = $HistoryHTML;
                    '_wfStatusChangeDate' = $SPInvoice.FieldValues._wfStatusChangeDate;
                    'AssignedTo1'         = $assignedToEmail;
                    'Business'            = $SPInvoice.FieldValues.Business;
                    'ConvertedAmount'     = $SPInvoice.FieldValues.ConvertedAmount;
                    'Currency'            = $SPInvoice.FieldValues.Currency;
                    'DAPSBank'            = $SPInvoice.FieldValues.Bank;
                    'EmailAlert'          = $SPInvoice.FieldValues.EmailAlert;
                    'InternalInvoice'     = $SPInvoice.FieldValues._InternalInvoice;
                    'InvoiceAmount'       = $SPInvoice.FieldValues.InvoiceAmount;
                    'InvoiceReceivedDate' = $SPInvoice.FieldValues.InvoiceReceivedDate;
                    'IsEmployee'          = $SPInvoice.FieldValues.IsEmployee;
                    'Priority'            = $SPInvoice.FieldValues.Priority;
                    'RAGDate'             = $SPInvoice.FieldValues.RAGDate;
                    'RAGStatus'           = $SPInvoice.FieldValues.RAGStatus;
                    'RefNo'               = "$($S.BU)|$($SPInvoice.Id)";
                    'StageNote'           = $SPInvoice.FieldValues.StageNote;
                    'StageRAGDate'        = $SPInvoice.FieldValues.StageRAGDate;
                    'StageRAGStatus'      = $SPInvoice.FieldValues.StageRAGStatus;
                    'Title'               = $SPInvoice.FieldValues.Title;
                    'VATAmount'           = $SPInvoice.FieldValues.VATAmount;
                    'wfSubStage'          = $SPInvoice.FieldValues.wfSubStage;
                    "DAPSBU"              = $SPInvoice.FieldValues.Business;
                    "DAPSInvoiceNo"       = $($SPInvoice.FieldValues.RefNo);
                    "DAPSNewItem"         = $false;
                    "DAPSUsed"            = $null;
                    "wfSubStage1"         = $($StageAction);
                } 
                
                $a = CheckApportionments -InvoiceData $thisItem

                #======================================================
                # of so new we just need to apply it to the O365 1tem. (i wonder if the field names are right ) 
                $a = Set-PnPListItem -Connection $destConnection -List "Invoices" -Identity $O365File.Id -Values $thisItem
                logActivity -Indent 1 -Type "4 Action" -Message "Updated the O365 Metadata"

                $a = Set-PnPListItem -Connection $thisConnection -List "Invoices" -Identity $SPInvoice.Id -Values @{"_wfLockTime" = $LockTime }
                logActivity -Indent 1 -Type "4 Action" -Message "set the Source item Lockdate in $($S.BU)"

            } else {
                logActivity -Indent 1 -Type "5 Warning" -Message "$($SPInvoice.FieldValues.FileLeafRef) is not a PDF"
            }
        } 
        #<#
        #NEW BIT (25th APRIL) TRASH THE CANCELLED ONES  (why test for lock time?)
        $SPDeadInvoices = Get-PnPListItem -Connection $thisConnection  -List "Invoices" | Where-Object { $_.FieldValues.wfSubStage -GT "A"  -and $_.FieldValues._wfLockTime -eq $null } 
        logActivity -Indent 1 -Type "2 Info" -Message "We found $($SPDeadInvoices.Count) items in the $($S.BU) site" 
        foreach ($DeadOne in $SPDeadInvoices) {
            logActivity -Indent 1 -Type "2 Info" -Message "Considering delete for $($SPInvoice.FieldValues._UserField2) - ($xStage) $($SPInvoice.FieldValues.wfSubStage)  $($SPInvoice.FieldValues.FileRef)  Assigned to $($assignedToEmail)"   -ForegroundColor Cyan
             
            #prepare the place its going to be killed in 
            $FileName = $DeadOne.FieldValues.FileLeafRef
            $DestFolderLocation = "/sites/PLRe-tDAPSApprovals/Invoices/$($S.BU)"
            $DestPath = $DestFolderLocation + "/" + $FileName

            #Get the condemned item? 
            $O365DeadFile = Get-PnPFile -Connection $destConnection -Url $DestPath -AsListItem   -ErrorAction SilentlyContinue

            #Check if File already exists 
            If ($O365DeadFile) {
                logActivity -Indent 1 -Type "4 Action" -Message "DELETE $($fileName) STILL exists in O365 will delete it (TO DO TRASH THE APPORTIONMENTS for $($O365DeadFile.Id) " 
                #DELETE THE ITEM AT THIS STAGE FROM 0365
                $d = Remove-PnPListItem -Connection $destConnection -List "Invoices" -Identity $O365DeadFile.Id -Recycle -Force
            }  else {
                logActivity -Indent 1 -Type "2 Info" -Message "FILE NOT THERE " 
            }
            $u2 = set-PnPListItem -Connection $thisConnection   -List "Invoices" -Identity $DeadOne.Id -Values @{"_wfLockTime" = $LockTime }
            logActivity -Indent 1 -Type "4 Action" -Message "UPDATE $($fileName) STILL exists in Source Site:$($S.BU) but lock time set to $($LockTime)" 
        }
        #>

    } else {
        logActivity -Indent 1 -Type "2 Info" -Message "DAPS Interface switched off for $($S.SiteURL)"
    }
    #WriteQuiet - only write a record IF the max error exceeds the limit $Script:LogControl.LogLevel 
    #Write always writes at lease one line  
    logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished in $($JobDuration) seconds" -logAction "WriteQuiet"
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




