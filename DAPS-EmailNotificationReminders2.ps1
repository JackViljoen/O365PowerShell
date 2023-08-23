Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "") 

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
#how old should invoices be to be included in the special users notification 
$NotificationThreshold = @{
    "==" = @{
            "Monday"   = 4
            "Tuesday"  = 2
            "Wednesday"= 2
            "Thursday" = 2
            "Friday"   = 500
            "Saturday" = 0
            "Sunday"   = 0
            }

    "=>" = @{
            "Monday"   = 4
            "Tuesday"  = 2
            "Wednesday"= 2
            "Thursday" = 2
            "Friday"   = 500
            "Saturday" = 0
            "Sunday"   = 0
            }
    "<=" = @{
            "Monday"   = 500
            "Tuesday"  = 500
            "Wednesday"= 500
            "Thursday" = 500
            "Friday"   = 500
            "Saturday" = 0
            "Sunday"   = 0
            }
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
   Param([parameter(position = 1)] $invoice, [parameter(position = 2)] $amberDays, [parameter(position = 2)] $redDays, [parameter(position = 4)] $Threshold)

   $Age = new-timespan -Start $invoice.FieldValues._wfStatusChangeDate -End (get-date)

   if($Age.Days -gt $Threshold){
        write-host "$($invoice.FieldValues.Title) Special case of a invoice we don't want notification of its not the weekly reminder age:$($Age.Days) threshold:$($Threshold)" -ForegroundColor Magenta
        return ""
   } else {
       if($Age.Days -gt $Script:MaxDays) {$Script:MaxDays = $Age.Days}

       #set the colour pairs for each rag 
       if($Age.Days -lt $amberDays){
            $b = "#8AC552"
            $c = "#000000"
       } else {
            if($Age.Days -lt $redDays){
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
       $Script:Lines++ 
       $Script:InvoiceIDs += ":$($invoice.id)"
       write-host "$($Age.Days) -[$($Threshold)days]  $($invoice.FieldValues.RefNo) : $($invoice.FieldValues.wfSubStage) - $($invoice.FieldValues.Title)"
       return $hline+$thisLine   #.replace("'",'"');
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

# Where is the site and the library to store the LOG into 
$LogSiteURL                = "https://pacificlife.sharepoint.com/sites/PLRe"
$LogLibName                = "wfHistoryEvents"

# Who are we savign the log as and conenct to the log site as them 
$logaccountName            = "svc_sp_sync@Pacificlifere.com" 
$logencrypted              = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
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
# TASK CODE HERE - CONNECT TO OPETATIONAL SYSTEM AND PROCESS, 
# NOTE: IT MAY BE THE SAME SITE / CREDENTIAL AS THE OPERATIONAL SITE BUT DONT REUSE THE LOG CREDENTIAL AND THE LOGCONNECTION
# CREATE NEW ONES AS THE CODE DEMANDS - THE LOG IS A BOLT ON IT SHOUDL NOT INTERFERE WITH THE EXISTING CODE 
#

$WorkingDay       = [string] (Get-Date).DayOfWeek ##  we need this as a string not a clever date thing else the lookup won't work 


if($WorkingDay -eq "Saturday" -or $WorkingDay -eq "Sunday"){
    write-host "Its a weekend nothign to do"
} else {
    $accountName  = "svc_sp_sync@Pacificlifere.com" 
    $encrypted    = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
    $credential   = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

    $SiteURL      = "https://pacificlife.sharepoint.com/sites/PLRe-tDapsApprovals"
    $Connection   = Connect-PnPOnline -Url $SiteURL -Credentials $credential -ReturnConnection

    #Now get the invoices at stages 1 to 4  with a send count of 0 
    $invoices = Get-PnPListItem -Connection $Connection -List "Invoices" `
    | Where-Object {$_.FieldValues.wfSubStage -ge "2.0" -and $_.FieldValues.wfSubStage -lt "5.0" -and $_.FieldValues.DAPSUsed -ne 1 -and $_.FieldValues.AssignedTo1 -ne $null -and $_.FieldValues.AssignedTo1.email -ne $null} `
    | Sort-Object  {$_.FieldValues.AssignedTo1, $_.FieldValues.wfSubStage, $_.FieldValues._wfStatusChangeDate } 

    if($invoices.Count -eq 0){
        logActivity -Indent 0 -Type "6 Error" -Message  "No Emails to be sent nothing to do WOT !!! cant be right go check"
    } else {

        $Specials = Get-PnPListItem -Connection $Connection -List "NotificationPreferences"
        $SpecialEmails  = @() 
        $SpecialEmails = $Specials | Select-Object {$_.FieldValues.User.email.toLower()} -Unique   #this is horrible 
        $SpecialEmailEmails  = @()
        foreach($EN in $SpecialEmails){ 
            if($EN.'$_.FieldValues.User.Email.toLower()' -gt "") {
                $SpecialEmailEmails += $EN.'$_.FieldValues.User.Email.toLower()'
            }
        }
         
        #do a map function to build JUST the emails we will need.  
        $emailNames   = @() 
        $emailNames   = $invoices | Select-Object {$_.FieldValues.AssignedTo1.email.toLower()} -Unique   #this is horrible 
        $EmailEmails  = @()
        foreach($EN in $emailNames){ 
            if($EN.'$_.FieldValues.AssignedTo1.email.toLower()' -gt "") {
                $EmailEmails += $EN.'$_.FieldValues.AssignedTo1.email.toLower()'
            }
        }
        $Diff = Compare-Object -ReferenceObject $EmailEmails -DifferenceObject $SpecialEmailEmails -IncludeEqual
 
        foreach($DI in $Diff){
            $AlertThreshold = $NotificationThreshold[$DI.SideIndicator][$WorkingDay]
            logActivity -Indent 0 -Type "3 Info" -Message  "$($DI.InputObject) $($DI.SideIndicator) has a threshold of $($AlertThreshold)" -ForegroundColor red
            $thisPersonsInvocies = $invoices |  Where-Object {($_.FieldValues.AssignedTo1.email).toLower() -eq $DI.InputObject} #get the emails for the current person 

            if($thisPersonsInvocies.Count -gt 0){

                logActivity -Indent 0 -Type "3 Info" -Message  "Working with  $($DI.InputObject) they have $($thisPersonsInvocies.Count) Items to process " -ForegroundColor Cyan
                #start the process for this person
            
                $Script:Lines = 0 
                $EmailBody = ""
                $Script:MaxDays = 0
                $Script:CurrentStage = ""
                $Script:InvoiceIDs = ""
                foreach($TPE in $thisPersonsInvocies){
                    #logActivity -Indent 0 -Type "1 Success" -Message  "Checking the status of $($TPE.FieldValues.BusRef)"
                    $personBU = $TPE.FieldValues.RefNo.split("|")[0] # don't like doing this more than once --- 
                    $EmailBody += produce-MailLine -invoice $TPE -amberDays 4 -redDays 7 -Threshold $AlertThreshold

                } 
                  
                if($Script:Lines -gt 0){
                    logActivity -Indent 0 -Type "4 Action" -Message  "SENDING EMAIL with $($Script:Lines) :items [IDs = $($Script:InvoiceIDs)] to $($DI.InputObject) "  -ForegroundColor Magenta
                    #who to send it to 
                    $RecipientName = (Get-Culture).TextInfo.ToTitleCase($DI.InputObject.split(".")[0].ToLower()) # A LOT of fuss to get a reliably capitalised first character
                    $EmailHeader  = "<h3>Daily DAPS Reminder Notification</h3><p>Hi $($RecipientName),</p>"

                    #now work out the body to send three emails to choose from- everybody: specials mon to thurs: Specials Friday 
                    if ($DI.SideIndicator -eq "<="){
                        write-host "Ordinary Joe"
                        $Nag = ""
                        if($Script:MaxDays.Days -gt 5) {$Nag = "<br/>Please be aware that one or more invoice(s) has been waiting for your approval for over $($Script:MaxDays.Days -1) days."}       
                        $EmailHeader += "<p>There are currently $($Script:Lines) invoice(s) for your attention:"+$Nag+"</p>"
                        $EmailHeader += "<p>You can use the individual links below OR you can see all of your invoices on the kanban page <a href='https://pacificlife.sharepoint.com/sites/PLRe-tDAPSApprovals/SitePages/Kanban.aspx'>here</a></p>"
                        $EmailHeader += "<table class='tablestyle'><tbody>"

                    } else {
                        if ($AlertThreshold -ne 500){
                            write-host "Special Mon - Thurs"
                            $EmailHeader += "<p>Please find the invoice(s) that have been marked for your attention over the previous working day. You can use the individual links in the table below to endorse or approve them. </p>"
                            $EmailHeader += "<p>NOTE: You will be sent a weekly reminder detailing all of your outstanding invoices on Friday. " 
                            $EmailHeader += "However, You can see then all at any time by visiting the kanban page <a href='https://pacificlife.sharepoint.com/sites/PLRe-tDAPSApprovals/SitePages/Kanban.aspx'>here</a></p>"
                        } else {
                            write-host "Special Friday"
                        }
                    }
                
                    $EmailFooter  = "</tbody></table><br/><h4>$($APCustomisations[$personBU].APSignoff)</h4>FYI the user guide for your business is here : $($APCustomisations[$personBU].APGuide)."
                
                    # MUNGE IT ALL TOGETHER!!!               
                    $Message = "$($EmailCss) $($EmailHeader) $($EmailBody) $($EmailFooter)"

                    #write-host $DI.InputObject $message -ForegroundColor Yellow
                    Write-host 
                    #FINALLY SEND THE DAMN THING!!! to 
                    $mail = smtp-SendEmail -RecipientEmails $DI.InputObject -CC $APCustomisations[$personBU].APAddress -BCC "tim.ellidge@pacificlifere.com" -From $APCustomisations[$personBU].APAddress -MessageText $Message -Subject "Daily DAPS Reminder" 

                } else {
                    write-host "Nothing to send to $($DI.InputObject)" -ForegroundColor Magenta
                }
                
                } else {
                write-host "$($DI.InputObject) [Alert threshold:$($AlertThreshold)days] has no invoices requiring notification" -ForegroundColor Green
            }
        }
        
    }
}
    




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