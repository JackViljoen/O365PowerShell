Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "") 

#  5.1 Pending Peoplesoft Coding
#5.0 Coding in PeopleSoft
#5.2 Peoplesoft Coding Issue
#6.0 Pending approval
#  

## UAT WITH BAD TIMING AND A WIDE RANCE OF ITEMS 


#=====================================================================================================================================
# Where is the Output Queue TAKE CARE !!! 
# LOCAL TEST = "C:\temp\InvoiceFiles\DAPS" ##LOCALTEST
# UAT        = "\\plre.pacificlife.net\plfuat2\InvoiceFiles\DAPS" ##UAT
# PROD       = Y:\" ##PRODUCTION
# RELAY      = "\\plreukpfil01v\London\Departments\Everyone\_TE\DAPS"

##DONT FORGET LINE 378

##=====================================================================================================================================
# Site detail including the Time Window - an hour to run the thing in ie every 5 min between the hours of 18 to 20 = "0/5 18-20 * * *"
$Sites = @(
            @{
                "active"      = $true
                "siteUrl"     = "https://pacificlife.sharepoint.com/sites/PLRe-UMeDAPS"
                "CRON"        = 8
                "TimeZone"    = "GMT Standard Time"
                "Contact"     = ""
                "FileDest"    = "Y:\"
                "BU"          = "UMe"
                "BUName"      = "UWME"
                "AccDtOffset" = 0
            } ,

            @{
                "active"      = $true
                "siteUrl"     = "https://pacificlife.sharepoint.com/sites/PLRe-AUDAPS"
                "CRON"        = 8
                "TimeZone"    = "AUS Eastern Standard Time"
                "Contact"     = ""
                "FileDest"    = "Y:\"
                "BU"          = "AU"
                "BUName"      = "AUST"
                "AccDtOffset" = 0
            },
                @{
                "active"      = $true
                "siteUrl"     = "https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS"
                "CRON"        = 11 
                "TimeZone"    = "GMT Standard Time"
                "Contact"     = ""
                "FileDest"    = "Y:\"
                "BU"          = "DC"
                "BUName"      = "EURO"
                "AccDtOffset" = 0
            }
        )

  # $TimeStuff = Get-TimeDifference "GMT Standard Time"
  # $TimeStuff = Get-TimeDifference "AUS Eastern Standard Time"
  # $TimeStuff = Get-TimeDifference "Pacific Standard Time"
##=====================================================================================================================================
# start simple send it the name of the field and then the Value to lookup
# if it finds nothing in the data map above it will return what it was sent
function Map() {
    Param([parameter(position = 0)] $set, [parameter(position = 1)] $d) 
    if(!$r){
        $r = ""
    } else {
        $r = $dataMapping[$set][$d]
        if($null -eq $r){
            $r = $d
        }
    }
    return $r
}


##=====================================================================================================================================
#  PEOPLESOFT DATA MAPPING 
$dataMapping = @{
                   "BU" = @{
                        "UMe" = "UWME"; 
                        "DC"  = "EURO"; 
                        "EU"  = "EURO"; 
                        "RE"  = "CANA"; 
                        "AU"  = "AUST";
                        "AS"  = "ASIA"
                    };
                    "PaymentType" = @{
                        "Wire Transfer"           ="WIR"
                        "Automated Clearing House"="ACH"
                        "xxx"                     ="CHK"
                        "Direct Debit"            ="DD"
                        "EFT"                     ="EFT"
                        "Manual Check"            ="CHK"
                        "Manual"                  ="MAN"
                        "BACS"                    ="EFT"
                    };
                    "Business" =@{
                        "PLIC Retro"    ="PLCND"
                        #AU
                        "PLRA"                          = "B04AU"
                        "PLRA (Intragroup Settlements)" = "B04AU"
                        "PLRA (NEOS/Investment)"        = "B04AU"
                        "PLRA (Regulatory/Tax/Rent)"    = "B04AU"
                        "UMA"                           = "G23AU"
                        "UMA (Regulatory/Tax/Rent)"     = "G23AU"
                        "UMA (Intragroup Settlements)"  = "G23AU"
                        "UMA (NEOS/Investment)"         = "G23AU"
                        "PLRL - CB"     ="C24CA"
                        #DC
                        "PLRS"          ="D09GB"
                        "PLRS-Europe"   ="D09GB"
                        "RGBM"          ="D05US"
                        "RIBM"          ="D25US"
                        "RIBM UK"       ="E33GB"
                        "RSBM"          ="D27US"
                        "RHBM"          ="D26US"
                        #RETRO
                        "PSCL"          ="PSCLC"
                        "RGBM Retro"    ="PLRBD"
                        "RGBM Retro a"  ="PLRBX"
                        "RIBM - CB"     ="C34CA"
                        #UME
                        "UMe Ltd"       ="PLUXX"
                        "UMTS Ltd"      ="E13GB"
                        "UMTSSB"        ="E22SG"
                        #ASIA
                        "RSSG"          ="A28SG"
                    };
                    "Bank" = @{
                        #UME
                        "HSBCT-E121"      = "HSBCT|E121"
                        "HSBCT-OP1"       = "HSBCT|OP1"
                        #AU
                        "ANZ Bank - PLRA" = "ANZ|OP1"
                        "ANZ Bank - UMA"  = "ANZ|UW1" 
                        #DC
                        "HSBCT-E091 GBP"      = "HSBCT|E091"        
                        "HSBCW-CUP8 USD"      = "HSBCW|CUP8"       
                        "HSBCT-E072 GBP"      = "HSBCT|E072"       
                        "HSBCW-CUP7 USD"      = "HSBCW|CUP7"       
                        "HSBCB-RIB2 USD"      = "HSBCB|RIB2"       
                        "HSBCB-RIB1 GBP"      = "HSBCB|RIB1"       
                        "HSBCB-RSB2 USD"      = "HSBCB|RSB2"       
                        "HSBCB-RSB1 GBP"      = "HSBCB|RSB1"       
                        "HSBCB-RGB2 GBP"      = "HSBCB|RGB2"       
                        "HSBCB-RGB1 USD"      = "HSBCB|RGB1"       
                        "HSBCB-RHB2 USD"      = "HSBCB|RHB2"       
                        "HSBCB-RHB1 GBP"      = "HSBCB|RHB1"
                    };
                     "REGION" = @{
                        "UMe" = "GBR"; 
                        "DC"  = "GBR"; 
                        "EU"  = "GBR"; 
                        "RE"  = "CAN"; 
                        "AU"  = "AUS";
                        "AS"  = "SGP"
                    };
                }

##=====================================================================================================================================
# General function that takes an object and a set of fields and preduces a CSV line for them 
function build-Record (){
    # a function so that im workgin with a big old string array rather than temp files
    # Much quicker and easier with a lower footfrint 
    Param([parameter(position = 0)] $Obj, [parameter(position = 1)] $Fields )

    #add the header record
    #$Script:LinesArray+= "'$($Fields.replace(",","','"))'"
    $fNames = $Fields.split(",").Trim()
     
    foreach($row in $Obj){
        $Result = ""
        foreach ($Nm in $fNames){
           $Result += """$($row[$Nm])""," 
        }
        $Script:LinesArray += $Result.TrimEnd(",")
    }
}

##=====================================================================================================================================
# WHERE'S THE MEAT YOU SAY? - IT'S HERE   
function Send-InvoicesFrom(){
    Param([parameter(position = 0)] $S)

    # so What sharepoint site are we interested in ? 
    $PingIndicator = ""
    $siteURL       = $S.SiteUrl
    $PSFileDest    = $S.FileDest
    $BU            = $S.BU
    $BUM = Map -set "BU" -d $BU
    
    $relativeRoot = $siteURL.ToString().Replace("https://pacificlife.sharepoint.com", "")
    $thisConnection = Connect-PnPOnline -Url $siteURL -Credentials $credential -ReturnConnection 
    logActivity -Indent 2 -Type "2 Info" -Message "Looking for invoices in $($BU) relative root $($relativeRoot) where it is $($TimeStuff.RemoteTime)" -ForegroundColor Green

    $Errors = ""
    $InvoiceArray = [System.Collections.ArrayList]::new()
    $Script:LinesArray=@()
    $i = 0 
    $lines = 0 

    ##start setting up the file data (it MAY not get written but best to prepare any way)
    $FileName      = "DAPS_$($BUM)_$(get-date -Format "yyyyMMdd_HHmm").csv"
    $Header        =  @{
    "RECORD_ID"    ="HDR"
    "FILE_NAME"    = $FileName
    "DTTM_CREATED" = Get-Date -format "MM/dd/yyyy HH:mm:ss"
    "DESCR100"     = "The peoplesoft file from $($BUM)"
    }

    Build-Record -Obj $Header -Fields "RECORD_ID,FILE_NAME,DTTM_CREATED,DESCR100"  

    #GET THE INVOICES IN THE STATUS RANGE
    $SPInvoices = Get-PnPListItem -Connection $thisConnection  -List "Invoices" | Where-Object {$_.FieldValues.wfSubStage -gt "5.0" -and $_.FieldValues.wfSubStage -lt "5.1"}
    $ALLApportionments = Get-PnPListItem -Connection $thisConnection  -List "InvoiceApportionments" -PageSize 1000 

    logActivity -Indent 3 -Type "1 Success" -Message "found  $($SPInvoices.Count) Invoices"

    foreach ($SPInvoice in $SPInvoices) {
           
        ##create an object with Mapping 

        $CorrectedDate = (Get-Date($SPInvoice.FieldValues.RAGDate)).AddSeconds($TimeStuff.RemoteUTCOffset + 3600)
        logActivity -Indent 3 -Type "3 Info" -Message "$($BUM) Invoice:$($SPInvoice.FieldValues.Title) coreected Date :$(Get-Date($CorrectedDate) -format "MM/dd/yyyy") Original date:$(Get-Date($SPInvoice.FieldValues.RAGDate) -format "MM/dd/yyyy")"

        $VendorRawCode = "0000000000$(($SPInvoice.FieldValues._Vendor).split("|")[1].Trim())" # make it long and pad it out
        $vendorCode    = $VendorRawCode.Remove(0, ($VendorRawCode.Length - 10)) # Now fix its length

        $thisItem =  @{
            "RECORD_ID"             = "INV";
            "VENDOR_SETID"          = "SHARE";
            "VNDR_LOC"              = " ";
            "ADDRESS_SEQ_NUM"       = 0;
            "DSCNT_AMT"             = 0;
            "ACCOUNTING_DT"         = Get-Date($AccountingDate)  -format "MM/dd/yyyy";
            'BUSINESS_UNIT'         = map -set "Business"  -d $SPInvoice.FieldValues.Business;
            'GROSS_AMT'             = $SPInvoice.FieldValues.InvoiceAmount;
            'INVOICE_DT'            = Get-Date($CorrectedDate) -format "MM/dd/yyyy";
            'VAT_ENTRD_AMT'         = $SPInvoice.FieldValues.VATAmount;
            'VENDOR_ID'             = $vendorCode
            'INVOICE_ID'            = $SPInvoice.FieldValues.RefNo;
            'CURRENCY_CD'           = $SPInvoice.FieldValues.Currency;
            'BANK_CODE'             = (map -set "Bank"       -d $SPInvoice.FieldValues.Bank).split("|")[0];
            'BANK_ACCT_KEY'         = (map -set "Bank"       -d $SPInvoice.FieldValues.Bank).split("|")[1];
            'PYMNT_METHOD'          = map -set "PaymentType" -d $SPInvoice.FieldValues._PaymentType;

            ## MY FIELDS TO PASS TO THE LINES
            'SP_ITEM_ID'            = $SPInvoice.FieldValues.ID;
            }
        #cut the PS Comment down to 30 characters 
        $str = $SPInvoice.FieldValues._UserField4
        if($str){
            $LineComment = $str.subString(0, [System.Math]::Min(30, $str.Length))  ;
        } else {
            $LineComment = ""
        }
       
        #Lets get a VAT % 
        $NetAmnt = $SPInvoice.FieldValues.InvoiceAmount - $SPInvoice.FieldValues.VATAmount
        $VatRatio = $NetAmnt / $SPInvoice.FieldValues.InvoiceAmount
        
        ## NOW THE LINE ITEMS OR APPORTIONMENTS
        ## GET IT BY DIVING RIGHT INTO THE FOLDER USING SERVER RELATIVE URL/  
        $Apportionments = $ALLApportionments | Where-Object {$_.FieldValues.BusRef -eq $SPInvoice.Id -or $_.FieldValues.BusRef -eq "$($SI.BU)|$($SPInvoice.Id)"}
               
        logActivity -Indent 4 -Type "1 Success" -Message  "Looked in $($relativeRoot)InvoiceApportionments/$($BU)-$($SPInvoice.Id) = found $($Apportionments.count) Apportionment records" -ForegroundColor Cyan    
        
        #now tidy them up and correct the amount
        $ApportionmentsArray = @()
       
        foreach($spApportionment in $Apportionments){
            #so the fields we populate depend upon the type or category of apportionment
            ## firstly lets build an or base record
            $thisApportionment =  @{
                "MERCHANDISE_AMT" = [math]::Round($spApportionment.FieldValues.APPAmount * $VatRatio,3);
                "APPCATEGORY"     = $spApportionment.FieldValues.APPCategory;
                "APPCODE"         = $spApportionment.FieldValues.APPCode;
            }
            $ApportionmentsArray+=$thisApportionment
        }

        ## IMPORTANT EXAMPLE IF ITS AN OBJECT USE THE $_[]
        $tidyArray = $ApportionmentsArray | Sort-Object -Property {$_["APPCATEGORY"],$_["APPCODE"], $_["MERCHANDISE_AMT"], $_["DESCR"]} -Unique 
        $Depts     = $tidyArray | Where-Object  {$_["APPCATEGORY"] -eq "Additional Cost Centre"} ##| Select-Object  {$_["APPCODE"]}  
        $Accounts  = $tidyArray | Where-Object  {$_["APPCATEGORY"] -eq "GL Code" } ##| Select-Object  {$_["APPCODE"]}  ## BE REALLY CAREDUL WITH RHE SPACE
        $MLOB      = $tidyArray | Where-Object  {$_["APPCATEGORY"] -eq "MLOB" } ##| Select-Object  {$_["APPCODE"]}  ## BE REALLY CAREDUL WITH RHE SPACE
        $PCODE     = $tidyArray | Where-Object  {$_["APPCATEGORY"] -eq "Project" } 
       
        #GO LIVE CHANGE REQUESTED BY AUS
        #did we get an MLOB? (australia) there can only be one its a compact little function 
        $thisMLOB = if($MLOB){$MLOB.APPCode} else {"MBXXX"}
        # END OF CHANGE 

        $NoDepts    = $Depts.APPCODE.Count
        $NoAccounts = $Accounts.APPCODE.Count

        $PSLinesArray = [System.Collections.ArrayList]::new()
        $ai = 1
        if(($NoDepts -gt 1 -and $NoAccounts -gt 1) -or $NoDepts -eq 0  -or $NoDepts -eq 0 ) { 
            logActivity -Indent 5 -Type "5 Warning" -Message  "ERROR RECORD $NoDepts :dept(s)    $NoAccounts :Account(s) returning to stage 5 " -ForegroundColor Red 
            If($BU -eq "DC"){
                $Values = @{
                    "PSVoucherNo"         = "NOT sumbitted to PS"; 
                    "PSReturnMessage"     = "Check apportionments in DAPS"; 
                    "wfSubStage"          = "5.0 Coding in PeopleSoft"; 
                    "_UserField2"         = "";
                }
                $a = set-PnpListItem -Connection $thisConnection -List "Invoices" -Identity $SPInvoice.Id -Values $Values
            }

        } else {
            logActivity -Indent 4 -Type "2 Info" -Message  "RECORD $NoDepts :depts    $NoAccounts :Accounts " -ForegroundColor Cyan
            if ($NoDepts -gt 1){
                foreach ($D in $Depts){
                    $thisApportionment =  @{
                        "RECORD_ID"       = "LIN";
                        "VOUCHER_LINE_NUM"= $ai;
                        "DESCR"           = $LineComment;
                        "MERCHANDISE_AMT" = $D.MERCHANDISE_AMT ;
                        "VAT_LN_ENT_AMT"  = "0";
                        "ACCOUNT"         = $Accounts.APPCODE ;
                        "OPERATING_UNIT"  = " " ;
                        "DEPTID"          = $D.APPCODE.split(" ")[0] ;
                        "PRODUCT"         = " " ; 
                        "FUND_CODE"       = "LBXXX"; ##LOB
                        "CLASS_FLD"       = map -set "REGION" -d $BU;     ## REGION BASED ON BU 
                        "PROGRAM_CODE"    = $thisMLOB; ## MLOB set earlier 
                        "AFFILIATE"       = " " ;
                        "CHARTFIELD1"     = " " ;
                    }
                    $a = $PSLinesArray.add($thisApportionment)
                    $ai++
                }
            } else { 
                If ($NoAccounts -gt 1){
                    foreach ($A in $Accounts){
                        $thisApportionment =  @{
                            "RECORD_ID"       = "LIN";
                            "VOUCHER_LINE_NUM"= $ai;
                            "DESCR"           = $LineComment;
                            "MERCHANDISE_AMT" = $A.MERCHANDISE_AMT ;
                            "VAT_LN_ENT_AMT"  = "0";
                            "ACCOUNT"         = $A.APPCODE;
                            "OPERATING_UNIT"  = " " ;
                            "DEPTID"          = $Depts.APPCODE.split(" ")[0]  ;
                            "PRODUCT"         = " " ; 
                            "FUND_CODE"       = "LBXXX"; ##LOB
                            "CLASS_FLD"       = map -set "REGION" -d $BU;       ## REGION BASED ON BU 
                            "PROGRAM_CODE"    = $thisMLOB ; ## MLOB
                            "AFFILIATE"       = " " ;
                            "CHARTFIELD1"     = " " ;
                        }
                        $a =  $PSLinesArray.add($thisApportionment)
                        $ai++
                    }
                
                } Else {
                    #its Just One Record one line 
                    $thisApportionment =  @{
                        "RECORD_ID"       = "LIN";
                        "VOUCHER_LINE_NUM"= $ai;
                        "DESCR"           = $LineComment;
                        "MERCHANDISE_AMT" = $Depts.MERCHANDISE_AMT ;
                        "VAT_LN_ENT_AMT"  = "0";
                        "ACCOUNT"         = $Accounts.APPCODE ;
                        "OPERATING_UNIT"  = " " ;
                        "DEPTID"          = $Depts.APPCODE.split(" ")[0] ;
                        "PRODUCT"         = " " ; 
                        "FUND_CODE"       = "LBXXX"; ##LOB
                        "CLASS_FLD"       = map -set "REGION" -d $BU;      ## REGION BASED ON BU 
                        "PROGRAM_CODE"    = $thisMLOB ; ## MLOB
                        "AFFILIATE"       = " " ;
                        "CHARTFIELD1"     = " " ;
                    }
                    $a = $PSLinesArray.add($thisApportionment)
                    $ai++
                }
            }
    

            #we are here coz there wasn't an error in the item so inclue the invoice in the data that will becoem the file
            $i++
            Build-Record -Obj $thisItem -Fields 'RECORD_ID,BUSINESS_UNIT,INVOICE_ID,INVOICE_DT,VENDOR_SETID,VENDOR_ID,VNDR_LOC,ADDRESS_SEQ_NUM,ACCOUNTING_DT,GROSS_AMT,VAT_ENTRD_AMT,DSCNT_AMT,
            CURRENCY_CD,BANK_CODE, BANK_ACCT_KEY, PYMNT_METHOD' 

            $lines += $PSLinesArray.Count                
            Build-Record -Obj $PSLinesArray -Fields 'RECORD_ID,VOUCHER_LINE_NUM,DESCR,MERCHANDISE_AMT,VAT_LN_ENT_AMT,ACCOUNT,DEPTID, OPERATING_UNIT,PRODUCT,FUND_CODE,CLASS_FLD,PROGRAM_CODE,AFFILIATE,CHARTFIELD1'

            # lets update the record 
            $Values = @{
                "PSVoucherNo"         = "Submitted to PS"; 
                "PSReturnMessage"     = ""; 
                "wfSubStage"          = "5.1 Pending Peoplesoft Coding"; 
                "_wfStatusChangeDate" = $DTTM_CREATED;
                "_UserField2"         = $FileName.split(".")[0];
            }
            $a = set-PnpListItem -Connection $thisConnection -List "Invoices" -Identity $SPInvoice.Id -Values $Values  
        }
    }
       
    ## save the output file but lets not bother if its empty.  ie there are 0 invoices that were good emough to add - so we don't need to do any more 
    if($i -gt 0){
        logActivity -Indent 3 -Type "2 info" -Message  "We have $($i) invoices to send over so writing the footer"
        $footer =  @{
            "RECORD_ID" ="END"
            "INVOICE_COUNT" = $i
            "INV_LINE_COUNT" = $lines   
        }
        Build-Record -Obj $Footer -Fields 'RECORD_ID,INVOICE_COUNT,INV_LINE_COUNT'

        #so write it locally first so we know what happened 
        $Script:LinesArray | Out-File -FilePath ".\1-APPLICATIONS\O365DAPS\Temp\$FileName"

        #TA DA !!!- NOW -  Write the data out to the PS Queue 
        $Script:LinesArray | Out-File -FilePath "$($PSFileDest)\$($FileName)" 
     
        # SO WE CAN TEST IT GOT THERE BY QUERYING IT
        $files = get-ChildItem -Path $PSFileDest | sort-object LastWriteTime -Descending | select-object -First 1 

        if($files.Name -eq $FileName){
            $PingIndicator = "$($BU)_" # PREFIX ON THE PING TO LET IT KNOW THIS WAS AN ACTION RUN
            logActivity -Indent 3 -Type "4 Action" -Message  "Produced file ($FileName) and sucessfully added to the PS Queue here [$($PSFileDest)]" -ForegroundColor green 
        } else {
            logActivity -Indent 3 -Type "6 Error" -Message  "ERROR $FileName did not make it to the PS Queue" 
        }
    } else {
        logActivity -Indent 3 -Type "3 Info" -Message  "FYI No File produced"
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

$PingPrefix = ""
$CronID = "$($JobName)_$($SI.BU)"

# Who are we working as then 
$accountName  = "svc_sp_sync@Pacificlifere.com" 
$encrypted    = Get-Content ".\8-VAULT\d7a18ffc-473c-4fb0-adc4-4a36a37a7402.txt" | ConvertTo-SecureString
$credential   = New-Object System.Management.Automation.PsCredential($accountName, $encrypted)

$today = (get-date).DayOfWeek
if($today -eq "Friday" -or  $today -eq "Saturday"){
    #woohoo nothing to do a weekend just them puny humans cant keep up with us machines... 
    logActivity -Indent 0 -Type "2 Info" -Message "Nothing to do today its a $today" -ForegroundColor Yellow
} else {
    $Hour = (Get-Date).Hour
    
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

    # now do each site 
    foreach ($SI in $Sites){
        if ($SI.active){
            if ($SI.CRON -eq $Hour) {
                $PingPrefix += "$($SI.BU)_"
                $PSFileDest = $SI.FileDest   ## get the site specific File location
                #Get some Time Zone information
                ##$TimeStuff = Get-TimeDifference "AUS Eastern Standard Time"
                $TimeStuff = Get-TimeDifference $SI.TimeZone
                $AccountingDate = (Get-date).AddDays($SI.AccDtOffset) ## set the date to the date the item will be ingested (ie tomorrow)
                logActivity -Indent 0 -Type "4 Action"-Message  "OK so its going to run for $($SI.siteUrl)... where it is $($TimeStuff.RemoteTime) with accountign date offset of $($SI.AccDtOffset)" -ForegroundColor Green
                #
                # POP THE CALL TO THE PAYLOAD IN HERE 
                #
                $a = Send-InvoicesFrom -S $SI
                
            } else {
                logActivity -Indent 0 -Type "2 Info" -Message  "SKIPPING Nothing to do for site $($SI.siteUrl) According to the CRON of $($SI.CRON)" -ForegroundColor yellow
            } 

        } else {
            logActivity -Indent 0 -Type "2 Info" -Message "INACTIVE $($SI.siteUrl) " -ForegroundColor magenta
        }
        logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished" -logAction "WriteQuiet"
    }
}

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
#logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished" -logAction "WriteQuiet"
#>

