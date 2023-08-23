Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "") 

#
# this module needs to find the items that are at stage 9, it then needs to get each item and prepare A set of JSON Objects. 
# The invoice itself
# The workflow history for the (with soem timing between entries)
# And the Apportionments (de Duplicated) 
# Slightly more complicated as the Apportionments are in a folder preceded by the # # BU and the invoice no, i.e. UMe345
# also Apportionwmetns master is really in the 365 world and NEED to be deleted from there once this is finished 
# and the workflow history is in FXXX but not insurmountable. 

function simple-Format {
    Param([parameter(position = 0)] $i)
    #for some data items its easy to default to a standad type of format as a string, ie  decimal precision 2dp and date format 
    $V = $i
    if ($i){
        switch($i.GetType()) {
            "String"    {$V = $i} 
            "Double"    {$V = ($i).ToString("N2")} 
            "DateTime"  {$V = [String] $(Get-Date($i) -Format "dd MMM yyyy HH:mm")}    
        }
    }
    return $V
}

function Set_ApportionmentHTML {
    #this is the specific variant for an Apportionment row in the collection. 
    Param([parameter(position = 0)] $O, [parameter(position = 1)] $N)

    $user = if($O._wfUser) {$O._wfUser.split("@")[0].replace("."," ")} else {'-'}
    
    $h = "<div class='$($N)row'>";
    $h +=  "<div class='C_APPCategory'>$($O.APPCategory)</div>"
    $h +=  "<div class='C_APPCode'>$($O.APPCode)</div>"
    $h +=  "<div class='C_APPDescription'>$($O.APPDescription)</div>"
    $h +=  "<div class='C_APPAmount'>$($O.APPAmount.ToString("N2"))</div>"
    $h +=  "<div class='C__wfUser'>$($user)</div>"
    $h +=  "<div class='C__wfTime'>$(Get-Date($O._wfTime) -Format "dd MMM yyyy HH:mm")</div>"
    $h += "</div>"   
    return $h
}

function Set_HistoryHTML {
    Param([parameter(position = 0)] $O, [parameter(position = 1)] $N)
    #this is the same for a worklow history item 
    
    $h = "<div class='$($N)row'>"
    $h += "<div class='C__wfTime'>$(Get-Date($O._wfTime) -Format "dd MMM yyyy HH:mm")</div>"
    if(!$O._wfUser){
        $h += "<div class='C__wfUser'>???</div>"
    } else {
        if($O._wfUser.indexOf("@") -gt -1){
            $h += "<div class='C__wfUser'>$($O._wfUser.split("@")[0].replace("."," "))</div>"
        } else {
            $h += "<div class='C__wfUser'>$($O._wfUser)</div>"
        }
    }

    $h += "<div class='C__wfPrevStage'>$($O._wfPrevStage)</div>"
    $h += "<div class='C__wfStageChange'>$($O._wfStageChange)</div>"
    $h += "<div class='C__wfStreamTime0'>$($O._wfStreamTime0)</div>"
    $h += "<div class='C__wfAction'>$($O._wfAction)</div>"     
    $h += "</div>"   
    return $h
}

function populateCoversheet {
    Param([parameter(position = 0)] $invoiceData, [parameter(position = 1)] $H_HTML, [parameter(position = 2)] $A_HTML)
    $CoverSheet = '<html lang="eng" xmlns:mso="urn:schemas-microsoft-com:office:office" xmlns:msdt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882">
    <head>
        <style>
            * {
                box-sizing: border-box;
                font-size: 12px;
            }

            body {
                font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
                font-size: 11px;
            }

            .header {
                background-color: #1b2952;
                width: 100%;
                color: whitesmoke;
                text-align: center;
                vertical-align: middle;
                padding: 15px 20px 5px 15px;
                font-size: 25px;
                font-family: Georgia, "Times New Roman", Times, serif;
            }

            .strapline {
                background-color: #1b2952;
                color: #a7eafb;
                width: 100%;
                text-align: center;
                padding: 10px 30px;
                margin: 0px;
                font-size: 18px;
            }

            #pagecontent {
                margin: 10px;
                padding: 5px;
                float: left;
                width: 700px;
            }

            .declaration {
                background-color: #d7d2cb;
                color: #1b2952;
                width: 100%;
                text-align: left;
                padding: 15px 10px;
                font-size: 15px;
                font-weight: 600;
                float: left;
            }

            .declaration span {
                font-size: 12px;
                font-style: italic;
                font-weight: 400;
                padding: 10px 0 0 0;
            }

            .pagetitle {
                width: 100%;
                float: left;
                padding: 5px 0;
                color: #e04303;
            }

            .sectiontitle {
                width: 100%;
                float: left;
                padding: 15px 0px 5px 0px;
                color: #3681b5;
            }

            .cardtitle {
                width: 100%;
                margin-top: 10px;
                color: #61574e;
            }

            .card100 {
                float: left;
                width: 100%;
                box-sizing: border-box;
            }

            .card50 {
                float: left;
                width: 50%;
            }

            .card33 {
                float: left;
                width: 33.333%;
            }

            .card25 {
                float: left;
                width: 25%;
            }

            .card75 {
                float: left;
                width: 75%;
            }

            .card66 {
                float: left;
                width: 66.666%;
            }

            .ts0 {
                font-size: 1.5rem;
                font-weight: 600;
            }

            .ts1 {
                font-size: 1.3rem;
                font-weight: 500;
            }

            .ts2 {
                font-size: 1.1rem;
                font-weight: 500;
            }

            .ts3 {
                font-size: 1.1rem;
                font-weight: 400;
            }

            .ts4 {
                font-size: 0.9rem;
                font-weight: 500;
            }

            .ts5 {
                font-size: 0.9rem;
                font-weight: 400;
            }

            .responseblock {
                width: 100%;
                float: left;
            }

            .rq {
                width: 35%;
                float: left;
                padding: 2px 5px 2px 5px;
                text-align: right;
                border-right: 1px solid silver;
            }

            .ra {
                width: 64%;
                float: left;
                padding: 2px 5px 2px 5px;
            }

            /* Hiden Class for the data conditionals  */
            .hiddenblock {
                display: none;
            }
                    
            .Apportionmentrow {
                width: 90%;
                float: left;
                border-top: 0.5px solid silver;
                padding-top: 5px;
                margin-left: 25px;
                 box-sizing: border-box;
            }

            .Apportionmentrow div {
                width: 16%;
                float: left;
                border-right:1px solid gainsborough;
                padding-left: 5px;
                font-size: 0.8rem;
                box-sizing: border-box;
            }

            .Apportionmentrow .C_APPCategory{
                width:16%;
            }
            .Apportionmentrow .C_APPCode{
                width:14%;
                text-align:right;
                padding-right:5px;
            }

            .Apportionmentrow .C_APPDescription{
                width:20%;
            }
            .Apportionmentrow .C_APPAmount {
                width: 12%;
                text-align: right;
                padding-right: 5px;
            }
            .Apportionmentrow .C_APPAmount::before {
                content: "£ ";
            }

            .Apportionmentrow .C__wfUser {
                width: 20%;
                text-align:right;
                padding-right:5px;
            }

            .Apportionmentrow .C__wfTime {
                width: 16%;
            }


            .Historyrow{
                width:90%;
                float:left;
                border-top:0.5px solid silver;
                padding-top:5px;
                margin-left:25px;
                box-sizing: border-box;

            }

            .Historyrow div{
                width:18%;
                float:left;
                border-right:1px solid gainsborough;
                padding-left:5px;
                font-size: 0.8rem;
                box-sizing: border-box;
            }

            .Historyrow .C__wfStreamTime0{
                width:8%;
                text-align: right;
                padding-right:5px;
            }

            .Historyrow .C__wfFormType, 
            .Historyrow .C__wfFormID,
            .Historyrow .C__wfLogComment,
            .Historyrow .C__wfStreamStatus
            {
                display:none;
            }




        </style>

        <title></title>
    </head>

    <body>
        <div class="header">[DAPSInstance] Archived Invoice</div>
        <div class="strapline">[Title]</div>
        <div id="pagecontent">
            <!-- this is the person block of HTML to be reproduced  also probaly cant use Flex?-->            
            <!-- the page has a lot of these blocks think about usign CSS to do the show hide -->

            <!-- thismay work to hide a block based on thw value  -->
            <div class="sectiontitle ts1">Key Dates</div>
            <div class="card50">
                <div class="responseblock">
                    <div class="rq ts4">Invoice Received</div>
                    <div class="ra ts5">[InvoiceReceivedDate]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Invoice Date</div>
                    <div class="ra ts5">[RAGDate]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Date of Payment</div>
                    <div class="ra ts5">[_wfStatusChangeDate]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Archive Date</div>
                    <div class="ra ts5">[DAPSArchiveDate]</div>
                </div>
            </div>
            <div class="card50">
                <div class="responseblock">
                    <div class="rq ts4">Invoice upload</div>
                    <div class="ra ts5">[Created]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">&nbsp;</div>
                    <div class="ra ts5">&nbsp;</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Total Days</div>
                    <div class="ra ts5">[DAPSDays]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4"></div>
                    <div class="ra ts5"></div>
                </div>
            </div>

            <!-- thismay work to hide a block based on thw value  -->
            <div class="sectiontitle ts1">Invoice Data</div>
            <div class="card50">
                <div class="responseblock">
                    <div class="rq ts4">Vendor</div>
                    <div class="ra ts5">[_Vendor]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Invoice ID</div>
                    <div class="ra ts5">[RefNo] ([_UserField2])</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Payment Type</div>
                    <div class="ra ts5">[_PaymentTypeName]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Bank Details</div>
                    <div class="ra ts5">[_PayeeSortCode][_PayeeAccountNo]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Is Employee</div>
                    <div class="ra ts5">[IsEmployee]</div>
                </div>
            </div>
            <div class="card50">
                <div class="responseblock">
                    <div class="rq ts4">Invoice Amount</div>
                    <div class="ra ts5">[InvoiceAmount]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Currency</div>
                    <div class="ra ts5">[_Currency]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Tax Amount</div>
                    <div class="ra ts5">[VATAmount]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">OPEX Value</div>
                    <div class="ra ts5">[ConvertedAmount]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">PS VouncherNo</div>
                    <div class="ra ts5">[PSVoucherNo]</div>
                </div>
            </div>

            <!-- thismay work to hide a block based on thw value  -->
            <div class="sectiontitle ts1">Approval & Release</div>
            <div class="card50">
                <div class="responseblock">
                    <div class="rq ts4">Business</div>
                    <div class="ra ts5">[_Business]</div>
                </div>
                <div class="responseblock">
                  <div class="rq ts4">Department</div>
                  <div class="ra ts5">[InvDept]</div>
               </div>
                <div class="responseblock">
                    <div class="rq ts4">Endorser</div>
                    <div class="ra ts5">[_EndorserEmail]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Authoriser</div>
                    <div class="ra ts5">[_AuthoriserEmail]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Secondary Authoriser</div>
                    <div class="ra ts5">[_SecAuthoriserEmail]</div>
                </div>
            </div>
            <div class="card50">
                <div class="responseblock">
                    <div class="rq ts4">Batch No</div>
                    <div class="ra ts5">[_PaymentBatch]</div>
                </div>
 
                <div class="responseblock">
                    <div class="rq ts4">Releaser</div>
                    <div class="ra ts5">[Releaser1Email]</div>
                </div>
                <div class="responseblock">
                    <div class="rq ts4">Secondary Releaser</div>
                    <div class="ra ts5">[Releaser2Email]</div>
                </div>
                <div class="responseblock">
                  <div class="rq ts4"></div>
                  <div class="ra ts5"></div>
              </div>
            </div>
            <div class="sectiontitle ts1">Apportionments & Allocations</div>
            <div class="card100">[apportionmentsHTML]</div>
            <div class="sectiontitle ts1">Workflow Events</div>
            <div class="card100">[historyHTML]</div>
        </div>
    </body>
    </html>'

    ## so thsi does the clever(ish) thing its still justy munging strigns
    ## it can be tidied up to include some formatting - OR CAN IT?  

    foreach ($k in $invoiceData.Keys) {
        $CoverSheet = $CoverSheet.replace("[$($k)]", (simple-Format $invoiceData[$k]) )
    } 
    
    $CoverSheet = $CoverSheet.replace("[apportionmentsHTML]", $A_HTML )
    $CoverSheet = $CoverSheet.replace("[historyHTML]", $H_HTML )   

    return $CoverSheet
}

#======================================================================================================================================
#
## LOG START (PREAMBLE) 
#
#======================================================================================================================================

$BatchSize = 100
$MasterSiteUrl = "https://pacificlife.sharepoint.com/sites/PLRe-tdapsApprovals"

##encode the site and the Business unit abrreviation here so we can step through them 
$Sites = @(
    #"DC|https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS"
    #"RE|https://pacificlife.sharepoint.com/sites/PLRe-REDAPS"
    #"UMe|https://pacificlife.sharepoint.com/sites/PLRe-UMEDAPS"
    "AU|https://pacificlife.sharepoint.com/sites/PLRe-AUDAPS"
)

$JobName = $MyInvocation.MyCommand.Name.split(".")[0]      #Get the Job name 
$r = $MyInvocation.MyCommand.Source                        #set up up for Relative Addressing everything is under O365PowerShell.
$rt = $r.Substring(0,$r.IndexOf("\O365PowerShell\") + 15)  # the path is everything up to and including the /O365PowerShell
Set-Location $rt

#SET UP LOGGING FOR THIS MODULE - LINK TO THE Code so it has the SCRIPT scope
.".\2-UTILITIES\SPLogger.ps1"
.".\2-UTILITIES\Utilities.ps1"
.".\3-CRON\CronSimple.ps1"

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

# FIRSTLY SET UP THE CREDENTIALS TO USE 
$accountName            = "svc_sp_sync@Pacificlifere.com" 
$encrypted              = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$credential             = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

#connect to the Master site too as we are goping to validate AND delete from there 
$MConnection            = Connect-PnPOnline -Url $MasterSiteUrl -Credentials $credential -ReturnConnection

$MApportionments = Get-PnPListItem -Connection $MConnection -List "Lists/InvoiceApportionments" -PageSize 2000
                              

# so lets make a start
forEach ($row in $sites) {
    $siteURL = $row.split("|")[1]
    $BUCode = $row.split("|")[0]
    $relativeRoot = $siteURL.ToString().Replace("https://pacificlife.sharepoint.com", "")
    Write-host "  "
    write-host "SCANNING =  $($siteURL)" -ForegroundColor Yellow    
   
    # CONNECT TO SHAREPOINT
    $thisConnection = Connect-PnPOnline -Url $siteURL -Credentials $credential -ReturnConnection

    Write-host "      Looking for completed invoices in " $thisWeb.Url " relative root" $relativeRoot -ForegroundColor Green
                                                                  
     ####  ###### #####   # #    # #    #  ####  #  ####  ######  ####  
    #    # #        #     # ##   # #    # #    # # #    # #      #      
    #      #####    #     # # #  # #    # #    # # #      #####   ####  
    #  ### #        #     # #  # # #    # #    # # #      #           # 
    #    # #        #     # #   ##  #  #  #    # # #    # #      #    # 
     ####  ######   #     # #    #   ##    ####  #  ####  ######  ####                                                                

    #PREPARE A QUERY THAT GETS THE INVOICES AT STAGE 9 
    $SPInvoices = Get-PnPListItem -Connection $thisConnection  -List "Invoices" | Where-Object {$_.FieldValues.wfSubStage -gt "9" -and $_.FieldValues._UserField1 -eq $null} | Select-Object -First $BatchSize #-ErrorAction SilentlyContinue -ErrorVariable ErrVar

    $AllHistory = $SPHistory = Get-PnPListItem -Connection $thisConnection -List "Lists/wfHistoryF" -PageSize 1000 

    Write-host "Found  $($SPInvoices.Count) Invoices"
    foreach ($SPInvoice in $SPInvoices) {
        if ($SPInvoice.FieldValues.File_x0020_Type -eq "pdf") {
            # IF ITS A PDF
            $xStage = $SPInvoice.FieldValues.wfSubStage.Substring(0, 1);
            #Prepare sme variables (nto sure they are needed here but here goes) 
            $historyArray = [System.Collections.ArrayList]::new()
            $ApportionmentsArray = [System.Collections.ArrayList]::new()
            Write-host "Trying $($SPInvoice.FieldValues._UserField2) - ($xStage) $($SPInvoice.FieldValues.wfSubStage)  $($SPInvoice.FieldValues.FileRef) "   -ForegroundColor Cyan
                                                                              
            #    # ######   #    # #  ####  #####  ####  #####  #   # 
            #    # #        #    # # #        #   #    # #    #  # #  
            #    # #####    ###### #  ####    #   #    # #    #   #   
            # ## # #        #    # #      #   #   #    # #####    #   
            ##  ## #        #    # # #    #   #   #    # #   #    #   
            #    # #        #    # #  ####    #    ####  #    #   #   
                                                           
            # in production this will be ID also now in folders so go figure that out 
            $Q2 = "<View><Query><Where><Eq><FieldRef Name='_wfFormID'/><Value Type='Text'>$($SPInvoice.ID)</Value></Eq></Where><OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy></Query></View>"
            $historyArray = [System.Collections.ArrayList]::new()
            $historyHTML = ""

            $folder = "$($relativeRoot)/Lists/wfHistoryF/F$($SPInvoice.ID)"
            write-host "         Folder path = $folder" -ForegroundColor Yellow

            #do the query need to remove the Duplicates... ##| Sort-Object {$_._wfTime} -Unique
            $SPHistory = $AllHistory | Where-Object {$_.FieldValues._wfFormID -eq  $SPInvoice.ID} | Sort-Object {$_["FieldValues._wfTime"]} -Descending  -ErrorAction SilentlyContinue -ErrorVariable ErrVar
            if (!$SPHistory) {               
                logActivity -Indent 0 -Type "6 Error" -Message "NO WORKLOW HISTORY FOUND for $($BUCode)|$($SPInvoice.ID)"
            } else {
                Write-host "$($SPHistory.count) wf History records" -ForegroundColor Yellow
                # lets reference this from the Invoice RAG date
                $compareDate = get-date($SPInvoice.FieldValues.RAGDate)
                foreach ($SPH in $SPHistory) {
                    #BUILD AN OBJECT FOR THIS items wfhistory record WITH THE METADATA
                    if($SPH.FieldValues._wfStageChange) {
                        $stageDuration = New-TimeSpan -Start $compareDate -End (get-date($SPH.FieldValues._wfTime))
                         $compareDate = get-date($SPH.FieldValues._wfTime) 
                        $thisHistoryRec = @{
                            '_wfFormType'     = $SPH.FieldValues._wfFormType;
                            '_wfFormID'       = $SPH.FieldValues._wfFormID;
                            '_wfTime'         = $SPH.FieldValues._wfTime;
                            '_wfUser'         = $SPH.FieldValues._wfUser;
                            '_wfAction'       = $SPH.FieldValues._wfAction;
                            '_wfStageChange'  = $SPH.FieldValues._wfStageChange;
                            '_wfStreamStatus' = $SPH.FieldValues._wfStreamStatus;
                            '_wfStreamTime0'  = $stageDuration.TotalDays.ToString("0.###");
                            '_wfLogComment'   = $SPH.FieldValues._wfLogComment;
                            '_wfPrevStage'    = $SPH.FieldValues._wfPrevStage;
                            '_wfLongComment'  = $SPH.FieldValues._wfLongComment;
                        }
                        #ADD IT TO AN ARRAY LIST (ITS MORE EFFICIENT WITH MEMORY)
                        $counter = $historyArray.Add($thisHistoryRec)
                    
                        $historyHTML += Set_HistoryHTML -O $thisHistoryRec -N "History"
                       
                    }
                }
            
                  
                  ##   #####  #####   ####  #####  ##### #  ####  #    # #    # ###### #    # #####  ####  
                 #  #  #    # #    # #    # #    #   #   # #    # ##   # ##  ## #      ##   #   #   #      
                #    # #    # #    # #    # #    #   #   # #    # # #  # # ## # #####  # #  #   #    ####  
                ###### #####  #####  #    # #####    #   # #    # #  # # #    # #      #  # #   #        # 
                #    # #      #      #    # #   #    #   # #    # #   ## #    # #      #   ##   #   #    # 
                #    # #      #       ####  #    #   #   #  ####  #    # #    # ###### #    #   #    ####  
                                                                                            
                # in production this will be ID also now in folders so go figure that out 
                $ApportionmentsArray = [System.Collections.ArrayList]::new()
                $apportionmentHTML = ""

                $folder = "$($relativeRoot)/Lists/InvoiceApportionments/$($BUCode)-$($SPInvoice.ID)"
                
                write-host "Folder path = $folder" -ForegroundColor Yellow
                #prepare the query 
                $Q3 = "<View><Query><Where><Eq><FieldRef Name='BusRef'/><Value Type='Text'>$($BUCode)|$($SPInvoice.ID)</Value></Eq></Where><OrderBy><FieldRef Name='APPCategory' /></OrderBy></Query></View>"

                $SPApportionments = Get-PnPListItem -Connection $thisConnection -List "InvoiceApportionments" -FolderServerRelativeUrl $folder -Query $Q3  #  -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                if (!$SPApportionments) {
                    logActivity -Indent 0 -Type "3 Info" -Message "NO Apportionment FOUND for $($BUCode)|$($SPInvoice.ID) trying WITHOUT the BU Prefix"
                    $Q3 = "<View><Query><Where><Eq><FieldRef Name='BusRef'/><Value Type='Text'>$($SPInvoice.ID)</Value></Eq></Where><OrderBy><FieldRef Name='APPCategory' /></OrderBy></Query></View>"
                    $SPApportionments = Get-PnPListItem -Connection $thisConnection -List "InvoiceApportionments" -FolderServerRelativeUrl $folder -Query $Q3  #  -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                    if (!$SPApportionments) {
                        logActivity -Indent 0 -Type "6 Error" -Message "NO Apportionment FOUND for $($BUCode)|$($SPInvoice.ID) NONE there either"
                        $app = $false
                    } else {
                        $app = $true
                    }
                } else {
                    $app = $true
                }

                if($app){
                    logActivity -Indent 0 -Type "2 Info" -Message  "$($SPApportionments.count) raw apportionment records" -ForegroundColor Yellow

                     ## IMPORTANT EXAMPLE IF ITS AN OBJECT USE THE $_[]
                    $tidyArray = $SPApportionments | Sort-Object -Property {$_.FieldValues.APPCategory, $_.FieldValues.APPCode, $_.FieldValues.APPDescription, $_.FieldValues._wfUser, $_.FieldValues.APPAmount } -Unique 

                    foreach ($a in $tidyArray) {
                        #BUILD AN OBJECT FOR THIS items wfhistory record WITH THE METADATA
                        $thisApportionment = @{
                            "Title"          = $a.FieldValues.Title;
                            "APPBU"          = $a.FieldValues.APPBU
                            "APPCategory"    = $a.FieldValues.APPCategory;
                            "APPCode"        = $a.FieldValues.APPCode
                            "APPDescription" = $a.FieldValues.APPDescription
                            "APPAmount"      = $a.FieldValues.APPAmount;
                            "BusRef"         = $a.FieldValues.BusRef
                            "_wfTime"        = $a.FieldValues._wfTime;
                            "_wfUser"        = $a.FieldValues._wfUser;
                            "APPSequence"    = $a.FieldValues.APPSequence;
                        }
                        #ADD IT TO AN ARRAY LIST (ITS MORE EFFICIENT WITH MEMORY)
                        $counter = $ApportionmentsArray.Add($thisApportionment)
                        $apportionmentHTML += Set_ApportionmentHTML -O $thisApportionment -N "Apportionment"
                    }

                                                 

                    # ============================================================
                    #BUILD AN OBJECT FOR THIS IMAGE WITH THE METADATA this is the core of the Archive record and the cover PDF 
                    $thisItem = @{
                        '_EndorserEmail'      = $SPInvoice.FieldValues._EndorserEmail;
                        '_AuthoriserEmail'    = $SPInvoice.FieldValues._AuthoriserEmail;
                        '_SecAuthoriserEmail' = $SPInvoice.FieldValues._SecAuthoriserEmail;
                        'Releaser1Email'      = $SPInvoice.FieldValues.AssignedTo1.Email;
                        'Releaser2Email'      = if ($SPInvoice.FieldValues.BusinessOwner){$SPInvoice.FieldValues.BusinessOwner.Email} else {"No second releaser"}
                        '_PayeeAccountNo'     = $SPInvoice.FieldValues._PayeeAccountNo;
                        '_PayeeSortCode'      = $SPInvoice.FieldValues._PayeeSortCode;
                        '_PaymentTypeName'    = $SPInvoice.FieldValues._PaymentType;
                        '_PaymentBatch'       = $SPInvoice.FieldValues._PaymentBatch;
                        'DAPSArchiveDate'     = $SPInvoice.FieldValues._SystemUpdateDate;
                        'ArchiveDate'         = Get-Date -Format  "yyyy-MM-dd"
                        '_UserField1'         = $SPInvoice.FieldValues._UserField1;
                        '_UserField2'         = "$($BUCode)-$($SPInvoice.ID)";
                        '_UserField3'         = $SPInvoice.FieldValues._UserField3;  #NEW?
                        '_UserField4'         = $SPInvoice.FieldValues._UserField4;  #NEW?
                        '_wfStatusChangeDate' = $SPInvoice.FieldValues._wfStatusChangeDate;
                        '_Business'           = $SPInvoice.FieldValues.Business;
                        '_Bank'               = $SPInvoice.FieldValues.Bank;         #NEW?
                        'ConvertedAmount'     = $SPInvoice.FieldValues.ConvertedAmount;
                        '_Currency'           = $SPInvoice.FieldValues.Currency;
                        'PSVoucherNo'         = $SPInvoice.FieldValues.PSVoucherNo;
                        'PSReturnMessage'     = $SPInvoice.FieldValues.PSReturnMessage;
                        'InvoiceAmount'       = $SPInvoice.FieldValues.InvoiceAmount;
                        'InvDept'             = $SPInvoice.FieldValues.InvDept;
                        '_Priority'           = $SPInvoice.FieldValues.Priority;
                        'RAGDate'             = $SPInvoice.FieldValues.RAGDate;
                        'InvoiceReceivedDate' = $SPInvoice.FieldValues.InvoiceReceivedDate;
                        '_RAGStatus'          = $SPInvoice.FieldValues.RAGStatus;
                        'StageRAGDate'        = $SPInvoice.FieldValues.StageRAGDate;
                        'Title'               = $SPInvoice.FieldValues.Title;
                        'VATAmount'           = $SPInvoice.FieldValues.VATAmount;
                        '_Vendor'             = $SPInvoice.FieldValues._Vendor;
                        'IsEmployee'          = $SPInvoice.FieldValues.IsEmployee;
                        'IsInternal'          = $SPInvoice.FieldValues._InternalInvoice;
                        'wfSubStage'          = $SPInvoice.FieldValues.wfSubStage;
                        'RefNo'               = $SPInvoice.FieldValues.RefNo;
                        'Created'             = $SPInvoice.FieldValues.Created;
                        'DAPSDays'            = ($SPInvoice.FieldValues._wfStatusChangeDate - $SPInvoice.FieldValues.Created).TotalDays;
                        'DAPSInstance'        = $BUCode;  
                        '_wfHistory'          = ($historyArray | ConvertTo-Json -Compress);
                        '_Apportionments'     = ($ApportionmentsArray | ConvertTo-Json -Compress);
                        '_wfStreamStatus'     = ($MApportionmentsArray | ConvertTo-Json -Compress);
                    } 
                   
                    #======================================================
                    # so NOW we need to add it to the Archive library (based on year? RAG Date )
                    $year = get-date($SPInvoice.FieldValues.RAGDate) -Format "yy"

                    ##So we may turn this into a PDF lets get the file we just copied over and add it to the cache as a file
                    ## we call it i get BUT its really a put into this location  
                    Get-PnPFile -Connection $thisConnection -Url $($SPInvoice.FieldValues.FileRef) -Path C:\JSP\PDFTooling\Input\DAPS -FileName $($SPInvoice.FieldValues.FileLeafRef) -AsFile -Force

                    ##so it seems OK lets have a look at the cover sheet 
                    $CoversheetHTML = populateCoversheet -invoiceData $thisItem -H_HTML $historyHTML -A_HTML $apportionmentHTML
                    ##save it tolocal disk as a file ready to prepended 
                    $CoversheetHTML | Set-Content -path c:\JSP\PDFTooling\Input\DAPS\coverpage.html -Force

                    try{
                        # turn that into a {PDF} i may bung in a stylesheet with the -s option 
                        C:\JSP\PDFTooling\Prince\engine\bin\prince.exe --page-size "A4" --page-margin "15mm" "C:\JSP\PDFTooling\Input\DAPS\coverpage.html" 
                        # use a different DLL to weld them together (CAT ?)
                        C:\JSP\PDFTooling\PDFtkServer\bin\pdftk.exe "c:\JSP\PDFTooling\Input\DAPS\coverpage.pdf" "C:\JSP\PDFTooling\Input\DAPS\$($SPInvoice.FieldValues.FileLeafRef)" cat output "C:\JSP\PDFTooling\Output\DAPS\$($SPInvoice.FieldValues.FileLeafRef)"
        
                        $IT = Add-PnPFile -Connection $thisConnection -Path "c:\JSP\PDFTooling\Output\DAPS\$($SPInvoice.FieldValues.FileLeafRef)" -Folder "InvoiceArchive$($year)" -Values $thisItem -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                        if($IT.UniqueId){
                            $ArchiveLabel = @{"_UserField1" = "Moved to InvoiceArchive$($year)"}
                             $TidyUp = $true
                             logActivity -Indent 0 -Type "4 Action" -Message "Moved $($SPInvoice.FieldValues.FileLeafRef) to InvoiceArchive$($year)"
                        } else {
                            $ArchiveLabel = @{"_UserField1" = "Failed to Move to InvoiceArchive$($year)"}
                            $TidyUp = $false
                            logActivity -Indent 0 -Type "6 Error" -Message "FAILED to move $($SPInvoice.FieldValues.FileLeafRef) toInvoiceArchive$($year)"
                        }
                        
                    } catch {
                        $ArchiveLabel = @{"_UserField1" = "PROBLEM Moving $($SPInvoice.FieldValues.FileLeafRef) to InvoiceArchive$($year)"}
                        $TidyUp = $false
                        logActivity -Indent 0 -Type "6 Error" -Message  "PROBLEM Moving $($SPInvoice.FieldValues.FileLeafRef) to InvoiceArchive$($year)"
                    }
                    $a = Set-PnPListItem -Connection $thisConnection  -List "Invoices" -Identity $SPInvoice.Id -Values $ArchiveLabel
                    ## goign to delete Invoice the wf-history and the apportionments 
                    #<#
                    if($TidyUp){
                        #Trash The WF History. 
                        if($SPHistory){
                            $foldername  = $SPHistory[0].FieldValues.FileDirRef
                            $endFolder = $foldername.split("/")[-1]
                            if((check-value $endFolder) -eq "String"){
                                $f= Remove-PnPFolder -Connection $thisConnection -Folder "Lists/wfHistoryF" -Name $endFolder -ErrorAction SilentlyContinue -ErrorVariable ErrVar -Recycle -Force
                            }
                            logActivity -Indent 0 -Type "4 Action" -Message  "Removed wfHistory Folder $($endFolder)"
                        }

                        #trash the Local Apportionments
                        if($SPApportionments){
                            $foldername  = $SPApportionments[0].FieldValues.FileDirRef
                            $endFolder = $foldername.split("/")[-1]
                            if((check-value $endFolder) -eq "String"){
                                $f= Remove-PnPFolder -Connection $thisConnection -Folder "Lists/InvoiceApportionments" -Name $endFolder -ErrorAction SilentlyContinue -ErrorVariable ErrVar -Recycle -Force
                            }
                            logActivity -Indent 0 -Type "4 Action" -Message  "Removed local Apportionemtns Folder $($endFolder)"
                        }

                        #Trash the Local invoice
                        $endItem = $SPInvoice.Id
                        if((check-value $endItem) -eq "Integer"){
                            $f = Remove-PnPListItem -Connection $thisConnection -List "Invoices" -Identity $endItem -Recycle -Force -ErrorAction SilentlyContinue -ErrorVariable ErrVar 
                            logActivity -Indent 0 -Type "4 Action" -Message  "Removed Local invoice $($endItem)"
                        }

                        #TO THINK ABOUT - STAGE 9 ITEMS IN THE KANBAN 
                      
                        #trash the Remote Apportionments
                        $DApportionments = $MApportionments | Where-Object {$_.FieldValues.BusRef -eq "$($BUCode)|$($SPInvoice.ID)"}  #  -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                        if($DApportionments.count -gt 0){
                            logActivity -Indent 0 -Type "4 Action" -Message  "Removing $($DApportionments.count) Hybrid Apportionment records for $($BUCode)|$($SPInvoice.ID)"
                            forEach ($DA in $DApportionments){
                                $remoteItem = $DA.Id
                                if((check-value $remoteItem) -eq "Integer"){
                                    $d = Remove-PnPListItem -Connection $MConnection -List "Lists/InvoiceApportionments" -Identity $remoteItem -Recycle -Force
                                }
                            }
                        } else {
                           logActivity -Indent 0 -Type "3 Info" -Message  "No matching apportionments for $($BUCode)|$($SPInvoice.ID) "
                        }
                    }
                    #>

                } else {
                    $ArchiveLabel = @{"_UserField1" = "MISSING APPORTIONMENTS for $($BUCode)|$($SPInvoice.ID)"}
                    $a = Set-PnPListItem -Connection $thisConnection  -List "Invoices" -Identity $SPInvoice.Id -Values $ArchiveLabel
                    logActivity -Indent 0 -Type "5 Warning" -Message  "MISSING APPORTIONMENTS for $($BUCode)|$($SPInvoice.ID)"
                }
            }              
        }        
    }
}






  