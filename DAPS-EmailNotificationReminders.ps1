

$APCustomisations =@{
    "UME" = @{
        "APAddress"="accounts@underwriteme.co.uk";            
        "APSignoff"="Thanks, Accounts team";
        "APGuide"="<a href='https://pacificlife.sharepoint.com/sites/PLRe-Wave/Shared%20Documents/DAPS%20Hybrid%20Training%20Manual%20UME.pdf'>Underwrite Me</a>"
        
        };
    "AU"  = @{
        "APAddress"="Accounts.PayableAUS@PacificLifere.com";  
        "APSignoff"="Thanks, Accounts team";
        "APGuide"="<a href='https://pacificlife.sharepoint.com/sites/PLRe-Wave/Shared Documents/Daps-Hybrid Training.pdf'>Australia</a>"
        };
    "DC"  = @{
        "APAddress"="accountspayableUK@pacificlifere.com";    
        "APSignoff"="Thanks, Treasury team";
        "APGuide"="<a href='https://pacificlife.sharepoint.com/sites/PLRe-Wave/Shared Documents/Daps-Hybrid Training.pdf'>Division Centre</a>"
        };
}

$EmailCss = "<style>"+
            " h4{margin:6px 0 8px 0;}"+
            " .iconclass {width:30px; text-align:center;}"+
            " .ragiconstyle {width:25px; height:25px; margin:3px; border-radius:7px; text-align:center;}"+
            " .invoicestyle {width:500px;text-align:left; padding-left:10px;}"+
            " .valuestyle {width:100px; text-align:right; padding-right:10px;}"+
            " .tablestyle {margin-left:40px;}"+
            "</style>"

function produce-MailLine (){
   Param([parameter(position = 2)] $invoice, [parameter(position = 3)] $amberDays, [parameter(position = 4)] $redDays)

   $Age = new-timespan -Start $invoice.FieldValues._wfStatusChangeDate -End (get-date)
   #$Age = new-timespan -Start $item.FieldValues._wfTime -End (get-date)
   if($Age -gt $Script:MaxDays) {$Script:MaxDays = $Age}

   #set the colour pairs for each rag 
   if($Age.TotalDays -lt $amberDays){
        $b = "#8AC552"
        $c = "#000000"
   } else {
        if($Age.TotalDays -lt $redDays){
            $b = "#FFB42F"
            $c = "#000000"
        } else {
            $b = "#E03C03"
            $c = "#FFFFFF"
        }
   }

   if($Script:CurrentStage -ne $invoice.Fieldvalues.wfSubStage){
      $Script:CurrentStage = $invoice.Fieldvalues.wfSubStage
      $hline = "</body></table> "+
               "<h4>Action / Stage : $($Script:CurrentStage)</h4>"+
               "<table class='tablestyle'> <thead> <tr> "+
               "<th class='ragstyle'>Days</th> "+
               "<th class='invoicestyle'>Invoice detail</th> "+
               "<th class='valuestyle'>Value</th> "+
               "<th class='linkstyle'>link</th> "+
               "</tr> </thead> <tbody>"
   } else {
      $hline =""
   }

 
   $thisLine = "<tr>"+
       "<td class='ragstyle'><div class='ragiconstyle' style='background-color:$($b); color:$($c)'>$($Age.Days)</div></td>"+
       "<td class='invoicestyle'>$($invoice.FieldValues.Title)</td>"+
       "<td class='valuestyle'>$(($invoice.FieldValues.InvoiceAmount).tostring("N2"))</td>"+
       "<td class='linkstyle'><a href='https://pacificlife.sharepoint.com/sites/PLRe-tDAPSApprovals/SitePages/PLRE-DAPS-Approval.aspx?InvoiceID=$($invoice.id)'>Open invoice</a></td>"+
   "</tr> "

   return $hline+$thisLine   #.replace("'",'"');
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

# Where is the site and the library to store the LOG into 
$LogSiteURL                = "https://pacificlife.sharepoint.com/sites/PLRe"
$LogLibName                = "wfHistoryEvents"

# Who are we savign the log as and conenct to the log site as them 
$logaccountName            = "svc_sp_sync@Pacificlifere.com" 
$logencrypted              = Get-Content ".\4-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$logcredential             = New-Object System.Management.Automation.PsCredential($logaccountName, $logencrypted)
$LogConnection             = Connect-PnPOnline -Url $LogSiteURL -Credentials $logcredential -ReturnConnection


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

#
# TASK CODE HERE - CONNECT TO OPETATIONAL SYSTEM AND PROCESS, 
# NOTE: IT MAY BE THE SAME SITE / CREDENTIAL AS THE OPERATIONAL SITE BUT DONT REUSE THE LOG CREDENTIAL AND THE LOGCONNECTION
# CREATE NEW ONES AS THE CODE DEMANDS - THE LOG IS A BOLT ON IT SHOUDL NOT INTERFERE WITH THE EXISTING CODE 
#
<#
$accountName  = "svc_sp_sync@Pacificlifere.com" 
$encrypted    = Get-Content ".\4-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$credential   = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

$SiteURL      = "https://pacificlife.sharepoint.com/sites/PLRe-tDapsApprovals"
$Connection   = Connect-PnPOnline -Url $SiteURL -Credentials $credential -ReturnConnection


#Now get the items with a send count of 0 
$emails = Get-PnPListItem -Connection $Connection -List "Lists/InvoiceNotifications" |  Where-Object {$_.FieldValues.NotificationType -eq "NewItem"}   | Sort-Object {$_.FieldValues.RecipientEmails, $_.FieldValues.wfSubStage, $_.FieldValues._wfTime } 

if($emails.Count -eq 0){
    logActivity -Indent 0 -Type "1 Success" -Message  "No Emails to be sent nothing to do"
} else {
    #lets get all of the live invoices so we can check if we still need to send the invite for each one
    $liveInvoices = Get-PnPListItem -Connection $Connection -List "Invoices" |  Where-Object {$_.FieldValues.wfSubStage -lt "5.0"}  
    #do a map function to get the RefNo and the wfsubstage
    $InvoiceStages = $liveInvoices | Select-Object {$_.FieldValues.RefNo, $_.FieldValues.wfSubStage} 

    #do a map funtiomn to build JUST the emails we will need 
    $emailNames    = $emails | Select-Object {$_.FieldValues.RecipientEmails.toLower()} -Unique   #this is horrible 

    $EmailEmails  = @()
    foreach($EN in $emailNames){ $EmailEmails += $EN.'$_.FieldValues.RecipientEmails.toLower()'}  #this just cleans up the collection to have a nicer name 
 
    foreach($AN in $EmailEmails){
        $thisPersonsEmails = $emails |  Where-Object {$_.FieldValues.RecipientEmails -eq $AN} #get the emails for the current person 

        logActivity -Indent 0 -Type "2 Info" -Message  "Working with  $($AN) they have $($thisPersonsEmails.Count) Items to process " -ForegroundColor Cyan
        #start the process for this person
            
        $Lines = 0 
        $EmailBody = ""
        $Script:MaxDays = 0
        $Script:CurrentStage = ""
        foreach($TPE in $thisPersonsEmails){
            #logActivity -Indent 0 -Type "1 Success" -Message  "Checking the status of $($TPE.FieldValues.BusRef)"
            $personBU = $TPE.FieldValues.DAPSBU # don't like doing this more than once --- 
            $Invoice = $liveInvoices |  Where-Object {$_.FieldValues.RefNo -eq $TPE.FieldValues.BusRef} 
            if($invoice){
                if($TPE.FieldValues.wfSubStage -eq $invoice.FieldValues.wfSubStage -and $invoice.FieldValues.AssignedTo1.Email.toLower() -eq $AN){
                    logActivity -Indent 0 -Type "4 Action" -Message  "STILL FRESH Adding to Email for $($AN) $($TPE.FieldValues.BusRef)   at stage $($TPE.FieldValues.wfSubStage)  INVOICE is at $($invoice.FieldValues.wfSubStage)" -ForegroundColor Green
                    #so build the line item for this email fragment 
                    $EmailBody += produce-MailLine -item $TPE -invoice $Invoice -amberDays 4 -redDays 7 
                    $Lines++ 
                    $currentSendCount = $TPE.FieldValues.NotifcationCount
                    if(!$currentSendCount){$currentSendCount = 1} else {$currentSendCount++} #test it exists and if not set it to 1 else bump it up  (Prescient) 
                   $a = Set-PnPListItem -Connection $Connection -List "Lists/InvoiceNotifications" -Identity $TPE.Id -Values @{"NotifcationCount" = $currentSendCount; "LastNotification" = (Get-Date -Format "o"); }
                } else {
                    logActivity -Indent 0 -Type "3 Info" -Message  "STALE Email for $($AN) $($TPE.FieldValues.BusRef) at stage $($TPE.FieldValues.wfSubStage)  INVOICE is now for $($invoice.FieldValues.AssignedTo1.Email) and at $($invoice.FieldValues.wfSubStage)" -ForegroundColor yellow
                    $a = Set-PnpListItem  -Connection $Connection -List "Lists/InvoiceNotifications" -Identity $TPE.Id -Values @{"NotificationType" = "Stale"}
                }
            } else {
                logActivity -Indent 0 -Type "3 Info" -Message  "Invoice $($TPE.FieldValues.BusRef) has moved on from this life - its no longer in the range we need " -ForegroundColor Red
                $a = Set-PnpListItem  -Connection $Connection -List "Lists/InvoiceNotifications" -Identity $TPE.Id -Values @{"NotificationType" = "Stale"}
            }
        }   
        # did that lot result in any content to send over to them ? 
        if($Lines -gt 0 -and $AN -gt ''){
#If($AN -eq "court.post@Pacificlifere.com"){
            logActivity -Indent 0 -Type "4 Action" -Message  "SENDING EMAIL TO $($AN)"  -ForegroundColor Magenta
            $Nag = ""
            if($Script:MaxDays.Days -gt 5) {$Nag = "<br/>One or more invoice(s) has been waiting for your approval for over $($Script:MaxDays.Days -1) days." }
            $RecipientName = (Get-Culture).TextInfo.ToTitleCase($AN.split(".")[0].ToLower()) # A LOT of fuss to get a reliably capitalised first character
                
            $EmailHeader  = "<h3>DAPS Reminder Notification</h3><p>Hi $($RecipientName),</p>"+
                            "<p>There are currently $($Lines) invoice(s) for your attention:"+$Nag+"</p>"+
                            "<p>You can use the individual links below OR you can see all of your invoices on the kanban page <a href='https://pacificlife.sharepoint.com/sites/PLRe-tDAPSApprovals/SitePages/Kanban.aspx'>here</a></p>"+
                            "<table class='tablestyle'><tbody>"
                
            $EmailFooter  = "</tbody></table>"+
                            "<h4>$($APCustomisations[$personBU].APSignoff)</h4>FYI the user guide for your business is here : $($APCustomisations[$personBU].APGuide)."
                
            # MUNGE IT ALL TOGETHER!!!               
            $Message = "$($EmailCss) $($EmailHeader) $($EmailBody) $($EmailFooter)"
                
            # FINALLY SEND THE DAMN THING!!!
            #TEST  $Mail = smtp-SendEmail -RecipientEmails "tim.ellidge@pacificlifere.com" -From $APCustomisations[$personBU].APAddress -MessageText "HI Tim" -Subject "DAPS Notification" 
            #mail = smtp-SendEmail -RecipientEmails $AN -CC $APCustomisations[$personBU].APAddress -BCC "tim.ellidge@pacificlifere.com" -From $APCustomisations[$personBU].APAddress -MessageText $Message -Subject "DAPS Notification" 

        } 
        
    }
#}
}

#>


# REPLACE ALL (logActivity -Indent 0 -Type "1 Success" -Message ) WITH (logActivity -Indent 0 -Type "2 Info" -Message) Then set the type and the indent pased on the nature of the Message. 
# TYPES ARE : "1 Success", "2 Info", "3 Info", "4 Action", "5 Warning", "6 Error"


#======================================================================================================================================
#
# TASK END HERE - CLOSE OUT THE LOG AND REGISTER THE PING
#
#======================================================================================================================================
$JobDuration   = ((get-date) -  $Script:LogControl.LogFirstCall).TotalSeconds # how long did the processing take
$ping          = @{"LastAlive" = get-date ; "Duration(s)" = $JobDuration ;}

#check if its the first one if so create a directory otherwise just save it  
if((test-path -Path ".\9-PINGS\$($JobName)") -eq $false){New-Item -Path ".\9-PINGS\$($JobName)" -ItemType directory}
$a             = $ping | Out-File -FilePath ".\9-PINGS\$($JobName)\$($JobName)$(Get-date -Format "yyMMdd-HH_mm").txt" 

#WriteQuiet - only write a record IF the max error exceeds the limit $Script:LogControl.LogLevel 
#Write always writes at lease one line  
logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished" -logAction "WriteQuiet"
#>