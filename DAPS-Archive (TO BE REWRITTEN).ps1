#
#  The archive processign based on DAPS Interface A & E 
#
# this module needs to find the items that are at stage 9, it then needs to get each s and prepare A set of JSON Objects. 
# The invoice itself
# The workflow history for the items and the Apportionments
# Slightly more complicated as the Apportionments are in a folder preceded by the # # BU and the invoice no, i.e. UMe345
# and the workflow history is in F354 but not insurmountable. 

#THIS GETS USED TO PRODUCE THE FILE NAME HASH
function get-hash([string]$textToHash) {
    $hasher = new-object System.Security.Cryptography.MD5CryptoServiceProvider
    $toHash = [System.Text.Encoding]::UTF8.GetBytes($textToHash)
    $hashByteArray = $hasher.ComputeHash($toHash)
    foreach($byte in $hashByteArray)
    {
      $result += "{0:X2}" -f $byte
    }
    return $result;
 }

# GET FILE Item associated with doc from URL
function simple-Format {
    Param([parameter(position = 0)] $i)
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

function Get-ItemFromUrl {
    Param([parameter(position = 0)] $web, [parameter(position = 1)] $list, [parameter(position = 2)] $url)
    
    #BUILD A QUERY
    $Q4 = "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$url</Value></Eq></Where><OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy></Query></View>"             
    $SPitem = Get-PnPListItem  -Web $web -List $list -FolderServerRelativeUrl $folder -Query $Q4  #  -ErrorAction SilentlyContinue -ErrorVariable ErrVar
    if ($SPitem) {
        return $SPitem
    }
    else {
        return $null
    }
}

function Set_ObjToHTML {
    Param([parameter(position = 0)] $O, [parameter(position = 1)] $N)
    
    $h = "<div class='$($N)row'>";
    foreach ($k in $O.Keys) {
        $h += "<div class='C_$k'>$($O[$k])</div>"
    } 
    $h += "</div>"   
    return $h
}

function Set_ApportionmentHTML {
    #this is the specific variant for the item
    Param([parameter(position = 0)] $O, [parameter(position = 1)] $N)
    
    $h = "<div class='$($N)row'>";
    $h += "<div class='C_APPCategory'>$($O.APPCategory)</div>"
    $h += "<div class='C_APPCode'>$($O.APPCode)</div>"
    $h += "<div class='C_APPDescription'>$($O.APPDescription)</div>"
    $h += "<div class='C_APPAmount'>$($O.APPAmount.ToString("N2"))</div>"
    $h += "<div class='C__wfUser'>$($O._wfUser.split("@")[0].replace("."," "))</div>"
    $h += "<div class='C__wfTime'>$(Get-Date($O._wfTime) -Format "dd MMM yyyy HH:mm")</div>"
    $h += "</div>"   
    return $h
}

function Set_HistoryHTML {
    Param([parameter(position = 0)] $O, [parameter(position = 1)] $N)

    
    $h = "<div class='$($N)row'>"
    $h += "<div class='C__wfTime'>$(Get-Date($O._wfTime) -Format "dd MMM yyyy HH:mm")</div>"
    if($null -eq $O._wfUser){
        $h += "<div class='C__wfUser'>???</div>"
    } else {
        $h += "<div class='C__wfUser'>$($O._wfUser.split("@")[0].replace("."," "))</div>"
    }

    $h += "<div class='C__wfPrevStage'>$($O._wfPrevStage)</div>"
    $h += "<div class='C__wfStageChange'>$($O._wfStageChange)</div>"
    $h += "<div class='C__wfStreamTime0'>$(($O._wfStreamTime0/24).ToString("N2"))</div>"
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
                color: #ffb958;
                width: 100%;
                text-align: center;
                padding: 10px 30px;
                margin: 0px;
                font-size: 15px;
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
                    <div class="rq ts4">GPB Value</div>
                    <div class="ra ts5">[ConvertedAmount]</div>
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

# FIRSTLY SET UP THE CREDENTIALS TO USE 
$Username = "plre\svc_SP_Admin"
$Password = ConvertTo-SecureString "IUyne07hxj,ds6GR£M*Wdys7ydd" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential($Username, $Password)

                              
##encode the site and the Business unit abrreviation here so we can step through them 
$Sites = @(
    ##"DC|http://sharepoint/DC/DCFINANCE/DAPS/",
    "UMe|http://sharepoint/UMe/DAPS/",
    "UMe|http://sharepoint/AU/AUFinance/DAPS/"
)

# so lets make a start
forEach ($row in $sites) {
    $siteURL = $row.split("|")[1]
    $BUCode = $row.split("|")[0]
    $relativeRoot = $siteURL.ToString().Replace("http://sharepoint", "")
    Write-host "  "
    write-host "      SCANNING =  $($siteURL)" -ForegroundColor Yellow    
   
    # CONNECT TO SHAREPOINT
    $thisConnection = Connect-PnPOnline -Url $siteURL -Credentials $myCreds -ReturnConnection 

    Write-host "      Looking for completed invoices in " $thisWeb.Url " relative root" $relativeRoot -ForegroundColor Green

    #Prepare sme variables (nto sure they are needed here but here goes) 
    $historyArray = [System.Collections.ArrayList]::new()
    $ApportionmentsArray = [System.Collections.ArrayList]::new()
                                                                           
     ####  ###### #####   # #    # #    #  ####  #  ####  ######  ####  
    #    # #        #     # ##   # #    # #    # # #    # #      #      
    #      #####    #     # # #  # #    # #    # # #      #####   ####  
    #  ### #        #     # #  # # #    # #    # # #      #           # 
    #    # #        #     # #   ##  #  #  #    # # #    # #      #    # 
     ####  ######   #     # #    #   ##    ####  #  ####  ######  ####                                                                

    #PREPARE A QUERY THAT GETS THE INVOICES AT STAGE 9 
    $Q1 = "<View><Query>
      <Where>
           <Eq>
               <FieldRef Name='wfSubStage' />
               <Value Type='Text'>9.0 Released</Value>
           </Eq>
      </Where>
   </Query></View>"

    $SPInvoices = Get-PnPListItem -Connection $thisConnection  -List "Invoices" -Query $Q1 #-ErrorAction SilentlyContinue -ErrorVariable ErrVar
    Write-host "         found  $($SPInvoices.Count) Invoices"
    foreach ($SPInvoice in $SPInvoices) {
        if ($SPInvoice.FieldValues.File_x0020_Type -eq "pdf") {
            # IF ITS A PDF
            $xStage = $SPInvoice.FieldValues.wfSubStage.Substring(0, 1);
            Write-host "         Trying $($SPInvoice.FieldValues._UserField2) - ($xStage) $($SPInvoice.FieldValues.wfSubStage)  $($SPInvoice.FieldValues.FileRef) "   -ForegroundColor Cyan

         #  try{
                                                                              
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

            $folder = "$($relativeRoot)Lists/wfHistoryF/F$($SPInvoice.ID)"
            write-host "         Folder path = $folder" -ForegroundColor Yellow

            #do the query 
            $SPHistory = Get-PnPListItem -Connection $thisConnection -Web $thisWeb -List "_wfHistoryF" -FolderServerRelativeUrl $folder -Query $Q2  #  -ErrorAction SilentlyContinue -ErrorVariable ErrVar
            if ($SPHistory) {
                Write-host "            $($SPHistory.count) wf History records" -ForegroundColor Yellow
                foreach ($SPH in $SPHistory) {
                    #BUILD AN OBJECT FOR THIS items wfhistory record WITH THE METADATA
                    if($SPH.FieldValues._wfStageChange) {
                        $thisHistoryRec = @{
                            '_wfFormType'     = $SPH.FieldValues._wfFormType;
                            '_wfFormID'       = $SPH.FieldValues._wfFormID;
                            '_wfTime'         = $SPH.FieldValues._wfTime;
                            '_wfUser'         = $SPH.FieldValues._wfUser;
                            '_wfAction'       = $SPH.FieldValues._wfAction;
                            '_wfStageChange'  = $SPH.FieldValues._wfStageChange;
                            '_wfStreamStatus' = $SPH.FieldValues._wfStreamStatus;
                            '_wfStreamTime0'  = $SPH.FieldValues._wfStreamTime0;
                            '_wfLogComment'   = $SPH.FieldValues._wfLogComment;
                            '_wfPrevStage'    = $SPH.FieldValues._wfPrevStage;
                            '_wfLongComment'  = $SPH.FieldValues._wfLongComment;
                        }
                        #ADD IT TO AN ARRAY LIST (ITS MORE EFFICIENT WITH MEMORY)
                        $counter = $historyArray.Add($thisHistoryRec)
                    
                        $historyHTML += Set_HistoryHTML -O $thisHistoryRec -N "History"
                    }
                }
            }
                  
              ##   #####  #####   ####  #####  ##### #  ####  #    # #    # ###### #    # #####  ####  
             #  #  #    # #    # #    # #    #   #   # #    # ##   # ##  ## #      ##   #   #   #      
            #    # #    # #    # #    # #    #   #   # #    # # #  # # ## # #####  # #  #   #    ####  
            ###### #####  #####  #    # #####    #   # #    # #  # # #    # #      #  # #   #        # 
            #    # #      #      #    # #   #    #   # #    # #   ## #    # #      #   ##   #   #    # 
            #    # #      #       ####  #    #   #   #  ####  #    # #    # ###### #    #   #    ####  
                                                                                            
            # in production this will be ID also now in folders so go figure that out 
            $Q3 = "<View><Query><Where><Eq><FieldRef Name='BusRef'/><Value Type='Text'>$($SPInvoice.ID)</Value></Eq></Where><OrderBy><FieldRef Name='APPCategory' /></OrderBy></Query></View>"
            $ApportionmentsArray = [System.Collections.ArrayList]::new()
            $apportionmentHTML = ""

            $folder = "$($relativeRoot)Lists/InvoiceApportionments/$($BUCode)-$($SPInvoice.ID)"
            write-host "         Folder path = $folder" -ForegroundColor Yellow

            #do the query 
            $SPApportionments = Get-PnPListItem -Connection $thisConnection -Web $thisWeb -List "InvoiceApportionments" -FolderServerRelativeUrl $folder -Query $Q3  #  -ErrorAction SilentlyContinue -ErrorVariable ErrVar
            if ($SPApportionments) {
                Write-host "            $($SPApportionments.count) wf History records" -ForegroundColor Yellow
                foreach ($a in $SPApportionments) {
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
            }
                                      
            ###### #    #   ##   # #       ####  
            #      ##  ##  #  #  # #      #      
            #####  # ## # #    # # #       ####  
            #      #    # ###### # #           # 
            #      #    # #    # # #      #    # 
            ###### #    # #    # # ######  ####  
                                      

            # so lets turn People into Email addesses as strings  (this is best way of ensuring users are valid after they leave)
            $Releaser1Email = $null


            if ($SPInvoice.FieldValues.AssignedTo1) {
                $AssignedTo = Get-PnPUser -Identity $SPInvoice.FieldValues.AssignedTo1.LookupId # WE WANT SOME DETAILS (EMAIL)
                $Releaser1Email = $AssignedTo.Email
            }
                
            $Releaser2Email = "No second releaser" 
            if ($SPInvoice.FieldValues.BusinessOwner) {
                $BusinessOwner = Get-PnPUser -Identity $SPInvoice.FieldValues.BusinessOwner.LookupId # WE WANT SOME DETAILS (EMAIL)
                $Releaser2Email = $BusinessOwner.Email
            }

            #>
                   

            # ============================================================
            #BUILD AN OBJECT FOR THIS IMAGE WITH THE METADATA this is the core of the data transfer 
            $thisItem = @{
                '_EndorserEmail'      = $SPInvoice.FieldValues._EndorserEmail;
                '_AuthoriserEmail'    = $SPInvoice.FieldValues._AuthoriserEmail;
                '_SecAuthoriserEmail' = $SPInvoice.FieldValues._SecAuthoriserEmail;
                'Releaser1Email'      = $Releaser1Email;
                'Releaser2Email'      = $Releaser2Email;
                '_PayeeAccountNo'     = $SPInvoice.FieldValues._PayeeAccountNo;
                '_PayeeSortCode'      = $SPInvoice.FieldValues._PayeeSortCode;
                '_PaymentTypeName'    = $SPInvoice.FieldValues._PaymentType;
                '_PaymentBatch'       = $SPInvoice.FieldValues._PaymentBatch;
                'DAPSArchiveDate'     = $SPInvoice.FieldValues._SystemUpdateDate;
                '_UserField1'         = $SPInvoice.FieldValues._UserField1;
                '_wfStatusChangeDate' = $SPInvoice.FieldValues._wfStatusChangeDate;
                '_Business'           = $SPInvoice.FieldValues.Business;
                'ConvertedAmount'     = $SPInvoice.FieldValues.ConvertedAmount;
                '_Currency'           = $SPInvoice.FieldValues.Currency;
                #'ID'                      = $SPInvoice.FieldValues.ID;                #CANT UPDATE ID .. oBVS
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
                '_UserField2'         = "$($BUCode)-$($SPInvoice.ID)";
                '_wfHistory'          = ($historyArray | ConvertTo-Json -Compress);
                '_Apportionments'     = ($ApportionmentsArray | ConvertTo-Json -Compress);
            } 
                   
            #======================================================
            # so NOW we need to add it to the New List (Archive)
             

            ##So we may turn this into a PDF lets get the file we just copied over and add it to the cache as a file 
            Get-PnPFile -Url $($SPInvoice.FieldValues.FileRef) -Path C:\bls\PDFTooling\Input -FileName $($SPInvoice.FieldValues.FileLeafRef) -AsFile -Force

            ##so it seems OK lets have a look at the cover sheet 
            $CoversheetHTML = populateCoversheet -invoiceData $thisItem -H_HTML $historyHTML -A_HTML $apportionmentHTML
            ##save it tolocal disk as a file ready to prepended 
            $CoversheetHTML | Set-Content -path c://bls/PDFTooling/input/coverpage.html -Force

            try{
                # turn that into a {PDF} i may bung in a stylesheet with the -s option 
                C:\bls\PDFTooling\Prince\engine\bin\prince.exe --page-size "A4" --page-margin "15mm" "c://bls/PDFTooling/Input/coverpage.html" 
                # use a different DLL to weld them together
                C:\bls\PDFTooling\PDFtkServer\bin\pdftk.exe "c://bls/PDFTooling/Input/coverpage.pdf" "c://bls/PDFTooling/Input/$($SPInvoice.FieldValues.FileLeafRef)" cat output "c://bls/PDFTooling/Output/$($SPInvoice.FieldValues.FileLeafRef)"
        
                $IT = Add-PnPFile -Path "c://bls/PDFTooling/Output/$($SPInvoice.FieldValues.FileLeafRef)" -Folder "InvoiceArchive" -Values $thisItem

            } catch {
                Write-host "That did not work check the invice PDF"
            }
            ## can thi sact as the queue to send to O365 (UME Instance) 
            ##
            ## Package up the Json object as a Data fiel to be used by the migration to O365
            ## copy over the file. This is a superset of Interface A

            #======================================================
            #prepare the data for putting out to disk
            $Outputdata = $thisItem | ConvertTo-Json

            #create a companion JSON file fOR each invoice (use the Hash code)
            $JSONFilename = get-hash($SPInvoice.FieldValues.FileLeafRef)
            write-host  "         " + $JSONFilename "  -  " $SPInvoice.FieldValues.FileLeafRef
            ## GET IT STORED TO DISK - this is magic too - I'm impressed by the Pipe thingie (phnar.... )  
            $Outputdata | out-file -Width 500 -FilePath "\\PLREUKPSPAPP01V\DAPSArchiveOutCache\$($JSONFilename).json"

            #======================================================
                    # we need to copy this item to the cache 
                    write-host "            saving to cache - $($SPInvoice.FieldValues.FileRef)"  -foregroundcolor green -NoNewline

                    Copy-Item "c:/bls/PDFTooling/Output/$($SPInvoice.FieldValues.FileLeafRef)" -Destination "\\PLREUKPSPAPP01V\DAPSArchiveOutCache\$($JSONFilename).pdf"


                  

            ##
            ## 
            if ($IT) {
                write-host "          Popped into archive and updated $($SPInvoice.FieldValues.FileRef) can be deleted"  -foregroundcolor red
                #~ REMEMBER WE WILL NEED TO DELETE THIS INVOICE NOT LOCK IT
                Remove-PnPListItem -List "Invoices" -Identity $SPInvoice.ID -Recycle -Force  
            }
            else {
                write-host "         did not find item in archive   $($SPInvoice.FieldValues.FileRef)"  -foregroundcolor red
            }
    #    }
    #   catch {
   #         write-host "         Problem with processing   $($SPInvoice.FieldValues.FileRef) will try again"  -foregroundcolor red
    #    }        
    }

}
}



  