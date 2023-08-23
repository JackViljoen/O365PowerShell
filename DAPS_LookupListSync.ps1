Param( [parameter(position = 1)] [Int32] $ID = 0, [parameter(position = 2) ] [string] $RuleName = "DAPS_LookupListSync") 
#& here is a pipe  | and a back slash  \   to copy save you looking around on the keyboard

function parseUrl () {
   Param( [parameter(position = 1)] $URL) #the URL of the path to the lib. or doc lib

   # you can throw a URL at this and it will return the path to the owning site (for the connection) and
   # the name of the library or list (it ignores end slashes and the Lists thing
   if ($URL[-1] -eq "/") {
      #chop off the last character if its a / as it may be because these URLs are typed in by a person
      $URL = $URL -replace “.$”
   }

   $L = $URL.split('/')[-1]  #Get the library name its the last item in the array
   $P = if ($URL.ToUpper().IndexOf('/LISTS/') -eq -1) { $URL.Length - $L.Length - 1 } else { $URL.ToUpper().IndexOf('/LISTS/') }
   $S = $URL.Substring(0, $P ) # dig it from the original URL so as to retain the case of of the original string
   $SiteDetail = @{
      SiteUrl = $S;
      LibName = $L;
      LibPath = "Lists/$($L)"
   }
   Return $SiteDetail
}


#======================================================================================================================================
## LOG START (PREAMBLE) 
#======================================================================================================================================
#SET RELATIVE ADDRESSING
$JobName = $MyInvocation.MyCommand.Name.split(".")[0]      #Get the Job name 
$r = $MyInvocation.MyCommand.Source                        #set up up for Relative Addressing everything is under O365PowerShell.
$rt = $r.Substring(0,$r.IndexOf("\O365PowerShell\") + 15)  # the path is everything up to and including the /O365PowerShell
Set-Location $rt

# LINK TO THE CODE MODULES
.".\2-UTILITIES\SPLogger.ps1"
.".\2-UTILITIES\Utilities.ps1"
.".\8-VAULT\UserManagement.ps1"

#SET UP & START THE LOG
$Log = @{
    "SiteURL"     = "https://pacificlife.sharepoint.com/sites/PLRe"
    "LibName"     = "wfHistoryEvents"
    "AccountName" = "svc_sp_sync@Pacificlifere.com" 
    "Contact"     = "tim.ellidge@Pacificlifere.com";
    "Level"       = "1 Success"; # LEVELS ARE : "1 Success", "2 Info", "3 Info", "4 Action", "5 Warning", "6 Error"
}

$L = start-Log -Log $Log -ID $ID -RuleName $RuleName
if(!$L) { write-host "STOPPING..." -foregroundColor Red; exit 1}  #CANT LOG SO STOP 

#======================================================================================================================================
# LOG SETUP END
#======================================================================================================================================
# Add any other Job Level params here 
$Job = @{
    "SiteURL"     = "https://pacificlife.sharepoint.com/sites/PLRe"
    "AccountName" = "svc_sp_sync@Pacificlifere.com" 
    "Credential"  = $null
}

$Cnx = Connect_Site -SiteUrl $Job.SiteURL -User $Job.AccountName
if(!$Cnx.Connection){
    logActivity -Indent 0 -Type "6 Error" -Message "JOB Finished - No Connection made for $($Job.AccountName) on site $($Job.SiteUrl)" -logAction "Write" 
    exit 1
} else {


    #& every list has many fields we dont want to copy items that sharepoint manages or may be internal or read only so lets ignore them we only want what's left
    $NoiseFields = @("Id", "ParentUniqueId", "Created", "Last_x0020_Modified", "Created_x0020_Date", "ContentTypeID", "Author", "Editor", "_HasCopyDestinations", "_CopySource", 
       "owshiddenversion", "WorkflowVersion", "_UIVersion", "_UIVersionString", "Attachments", "_ModerationStatus", "_ModerationComments", "Edit", 
       "LinkTitleNoMenu", "LinkTitle", "LinkTitle2", "SelectTitle", "InstanceID", "Order", "GUID", "WorkflowInstanceID", "FileRef", "FileDirRef", "FSObjType", "SortBehavior", 
       "PermMask", "UniqueId", "SyncClientId", "ProgId", "ScopeId", "File_x0020_Type", "HTML_x0020_File_x0020_type",
       "_EditMenuTableStart", "_EditMenuTableStart2", "_EditMenuTableEnd", "DocIcon", "ServerUrl", "EncodedAbsUrl", "MetaInfo", "_Level", "_IsCurrentVersion", 
       "ItemChildCount", "FolderChildCount", "Restricted", "OriginatorId", "NoExecute", "ContentVersion", "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime",
       "_ComplianceTagUserId", "AccessPolicy", "_VirusStatus", "_VirusVendorID", "_VirusInfo", "AppAuthor", "AppEditor", "SMTotalSize", "SMLastModifiedDate", "SMTotalFileStreamSize",
       "SMTotalFileCount", "ComplianceAssetId", "CalculatedTitle", "PrincipalCount", "FileLeafRef", "LinkFilenameNoMenu", "LinkFilename", "LinkFilename2", "BaseName", 
       "_IsRecord", "_CommentFlags", "_CommentCount" , "_SourceListGUID", "_SourceItemGUID", "_SourceItemModified")

    #include the fields this interface uses to ignore as they must stay fresh, ie they ! 

    #& where are we getting the data from ? Note only the live sites... 
    $Sites = @(
       "https://pacificlife.sharepoint.com/sites/PLRe-AUDAPS"
       "https://pacificlife.sharepoint.com/sites/PLRe-UMEDAPS"
       "https://pacificlife.sharepoint.com/sites/PLRe-DCDAPS"
    )
    $sourceListName = "Lists/Apportionments"

    #& where is it going too? 
    $FollowerLists = "sites/PLRe-tdapsApprovals/Lists/Apportionments"

    $CopyData = $true ## lets assume we will copy stuff out 

    foreach ($SourceSiteUrl in $Sites){
        $sourceConnection = Connect-PnPOnline -Url $sourceSiteURL -ReturnConnection -Credentials $Cnx.Credential -ErrorAction SilentlyContinue -ErrorVariable ErrVar
        $LastRun = (get-date).AddYears(-10)

        if ($Errvar[0].Count -eq 0) {
           logActivity -Indent 0 -Type "1 Success" -Message  "Connected to the Source $($sourceConnection.Url)"
           $SourceList = Get-PnPList -Connection $sourceConnection -Identity $sourceListName -ThrowExceptionIfListNotFound -ErrorAction SilentlyContinue -ErrorVariable ErrVar
           if ($Errvar[0].Count -eq 0) {
                ##$ListLastDelete = $InfluencerList.LastItemDeletedDate
                $ListLastUpdate = $SourceList.LastItemModifiedDate
                $sourceListGUID = $SourceList.Id

                if ($ListLastUpdate -gt $LastRun) {
                    #Build a data structure to hold my changes just IDs for the delete but the add and update are heavier
                    # so use an array list as its more performant when it gets big and ill be addign to it - as i acrue changes from the masterlist
                    $ItemData = New-Object -TypeName 'System.Collections.ArrayList';
          
                    $payLoad = @{}
                    # Will pop them in the payload later thought i suppose it could have gone in there right away

                    #finally get the items from the source list i think i need them all to start with as i need to be able to identify the deletes (NOTE ONLY GET APPROVED ONES) 
                    $sourceItems = Get-PnPlistItem -Connection $sourceConnection -List  $sourceListName  -PageSize 500 | Where-Object {$_.FieldValues._ModerationStatus -eq 0}
                    ## NOT MUCH USE AS I CANT GET TYPE
                    logActivity -Indent 3 -Type "2 Info" -Message  "There are $($sourceItems.count) source records to process - maybe something to delete"
                    #TODO what happens with an empty List?  is that valid? It may / shoudl default to deletes?
                    if ($sourceItems.count -gt 0) {
                        #so we have at least 1 item what fields does it contain ignoring the ones i know we don't need?
                        $ListFieldskeys = @()
                        foreach ($IKey in $sourceItems[0].FieldValues.keys) {
                            if ($NoiseFields -notcontains $IKey) {
                                $ListFieldskeys += $IKey
                                try {
                                    $keyType = $sourceItems[0].FieldValues[$IKey].GetType()
                                }
                                catch {
                                    $keyType = "null or unknown"
                                }
                            }
                        }
                        logActivity -Indent 3 -Type "3 Info" -Message "we have $($ListFieldskeys.Length) fields [ $($ListFieldskeys -join "-") ]"


                        #~ No need to worry about deletes here we can delegate that to the follower list just send it the GUIDs to not delete
                        ## lets include it all in the payload and let the downstream lists sort it out

                        foreach ($II in $sourceItems) {
                            # logActivity -Indent 5 -Type "2 Info" -Message "ITEM $($II.Fieldvalues.ID)  $($II.FieldValues.Modified)"
                            $itemValues = @{}
                            $itemValues["_SourceItemGUID"] = $II.Fieldvalues.GUID
                            $itemValues["_SourceListGUID"] = $sourceListGUID
                            $itemValues["_SourceItemModified"] = get-date($II.FieldValues["Modified"]) -Format "yyyy-MM-ddTHH:mm:ssZ" 
                            foreach ($key in $ListFieldskeys ) {
                                ## we can't just do an allocation of the value as dates and lookups and other fields are a bit picky
                                ## we an extend this to cope with other things / field Types
                                if ($null -eq $II.FieldValues[$key]) {

                                }
                                else {
                                    $FieldType = "Unknown"
                                    #write-host "$key : Null"
                                    try {
                                        $FieldType = $II.FieldValues[$key].GetType()
                                        #write-host "$key : $FieldType"
                                        switch ($FieldType) {
                                            #
                                            ##FIELD TYPE OPTIONS - THESE ARE THE MAGIC MOVES - AS WE CANT COPY IN JUST VALUES FOR ALL FIELD TYPES
                                            #
                                                                           
                                            #LOOKUP 
                                            "Microsoft.SharePoint.Client.FieldLookupValue" { # lots of work to do for Lookups
                                                        $newKey  = $key + "_x007e_"  # append the _x007e_ it with the new thing 
                                                        $newKey = $newKey.subString(0, [System.Math]::Min(32, $newKey.Length)) # keep its length down ;
                                                        $itemValues[$newKey] = "$($II.FieldValues[$key].LookupValue)"  # embed the old and new names
                                                    } 
                                    
                                            #USER
                                            "Microsoft.SharePoint.Client.FieldUserValue" { $itemValues[$key] = $II.FieldValues[$key].Email }
                                    
                                            #DATE
                                            "datetime" { $itemValues[$key] = get-date($II.FieldValues[$key]) -Format "yyyy-MM-ddTHH:mm:ssZ" }
                                    
                                            #Everything Else
                                            default { $itemValues[$key] = $II.FieldValues[$key] }
                                        }
                                    } catch {
                                        $FieldType = "Unknown"
                                        #write-host "$key : unknown"
                                    }
                                }

                            }
                            $x = $ItemData.add(@{"values" = $itemValues })
                   
                        }
                        #do this as two seperate lists because we may be comparing GUIDS into more than one follower so it lets us compare more easilly
                
                        $payLoad['ItemData'] = $ItemData

                        logActivity -Indent 2 -Type "3 Info" -Message "We have built the payload object to send to the follower lists"
                    } else {
                        logActivity -Indent 3 -Type "2 Info" -Message "nothing to do on list:$($sourceListName) its empty"
                        $CopyData = $false
                    }
                } else {
                    logActivity -Indent 2 -Type "2 Info" -Message "Nothing to do as its run since the last update "
                    $CopyData = $false
                }
               
        #*##########################################################
        #* PART 2 Sending out the changes to the followers
        #*##########################################################

                #$CopyData = $false
                ## now do the distribution process we have a Payload its the the data we need from the sourcel list
                ## but there may be nothing to do so lets check
                if ($CopyData) {

                    #SOURCE COMPARISON DATA
                    $SourceListGUID = $payLoad.ItemData[0].Values._SourceListGUID
                    $itemGUIDs     = New-Object -TypeName 'System.Collections.ArrayList';  #EXISTS DELETE NEW
                    $itemGUIDDates = New-Object -TypeName 'System.Collections.ArrayList';  #UPDATES
                    $tot=0
                    foreach($thing in $Payload.ItemData){
                        $a = $itemGUIDs.add($thing.Values._SourceItemGUID)
                        $a = $itemGUIDDates.add("$($thing.Values._SourceItemGUID) $($thing.Values._SourceItemModified)")
                        $tot++
                    }

                    logActivity -Indent 2 -Type "2 Info" -Message "Now pushing the changes out to the followers"
            
                    $U = ""
                    $FollowerListItems = $FollowerLists.split("|").Trim()
                    foreach ($f in $FollowerListItems) {
                        if ($f.length -gt 0) {
                            #useful test if the user has ended the string with a ,
                            logActivity -Indent 3 -Type "2 Info" -Message "Checking then Pushing changes out to $f"
                            $p = parseurl -URL $f
                            #check the site url first
                            if ($p.SiteURL -ne $U) {
                                #if we have gone to a new site them remake the connection
                                $U = $p.SiteURL
                                $UpdateConnection = Connect-PnPOnline -Url "https://pacificlife.sharepoint.com/$($U)" -ReturnConnection  -Credentials $credential -ErrorAction SilentlyContinue -ErrorVariable Errvar
                            }

                            #can we get the list ?
                            $l = Get-PnPList -Connection $UpdateConnection -Identity $p.LibPath
                            if ($null -ne $l) {
                                #so here is where we get visibility of the list (that we know exists for the first time and here is where we need to see if it has the key field
                                $a = Get-PnPField -Connection $UpdateConnection -List $p.LibPath -Identity "_SourceItemGUID" -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                                if ($Errvar[0].Count -ne 0) {
                                    logActivity -Indent 3 -Type "5 Warning" -Message  "Follower list $($p.LibPath) does not yet have the 'SourceItemGUID' foriegn key) - adding it now - be patient) "
                                    $newField = Add-PnPField -Connection $UpdateConnection -List $p.LibPath -DisplayName "_SourceItemGUID" -InternalName "_SourceItemGUID" -Type Text -Group "PLRe Columns"
                                    $newField = Add-PnPField -Connection $UpdateConnection -List $p.LibPath -DisplayName "_SourceItemModified" -InternalName "_SourceItemModified" -Type Text -Group "PLRe Columns"
                                    $newField = Add-PnPField -Connection $UpdateConnection -List $p.LibPath -DisplayName "_SourceListGUID" -InternalName "_SourceListGUID" -Type Text -Group "PLRe Columns"         
                                } 
                                # Carry On now the fields exist  
                                # So the rule is we need to remove items with No source ITEM GUID irrespective of the list they come from 
                                logActivity -Indent 4 -Type "2 Info" -Message  "Checking for items with No GUID"
                                $InvalidLookupFollower = Get-PnPListItem -Connection $UpdateConnection -List $p.LibPath  | Where-Object { $_.FieldValues._SourceItemGUID -eq $null -or $_.FieldValues._SourceItemGUID -eq "" } 
                                if($InvalidLookupFollower){
                                    #we need to delete the invalid ones
                                    logActivity -Indent 4 -Type "4 Action" -Message "We have $($InvalidLookupFollower.Count) items with no GUID to delete" -ForegroundColor green 
                                    # set up a batch
                                    $DeleteTransactions = New-PnPBatch -Connection $UpdateConnection
                                    foreach($ILF in $InvalidLookupFollower){
                                        $deadThing = Remove-PnPListItem -Connection $UpdateConnection -List $p.LibPath -Identity $ILF.Id -Force -Recycle # -Batch $DeleteTransactions
                                    }
                                    #write-host "   Submitting batch to delete items" -ForegroundColor green 
                                    #Invoke-PnPBatch -Batch $DeleteTransactions -Connection $UpdateConnection
                                } else {
                                    logActivity -Indent 4 -Type "2 Info" -Message  "   All items have a GUID for Sync tracking" -ForegroundColor green 
                                }

                                #FOLLOWER COMPARISON DATA - FOR THIS LIST 
                                $FollowerItems = Get-pnpListItem -connection $UpdateConnection -List $p.LibPath -PageSize 500 | Where-Object {$_.FieldValues._SourceListGUID -eq $SourceListGUID}

                                #what items are in the follower list? ie the value of the SourceItemGUIDs and the Dates so we can so the Compares for add delete and update
                                $FollowerItemGUIDs     = New-Object -TypeName 'System.Collections.ArrayList';
                                $FollowerItemGUIDDates = New-Object -TypeName 'System.Collections.ArrayList';
                                ForEach ($FI in $FollowerItems) {
                                   $a = $FollowerItemGUIDs.Add($FI.FieldValues._SourceItemGUID)
                                   $a = $FollowerItemGUIDDates.Add("$($FI.FieldValues._SourceItemGUID) $(Get-Date($FI.FieldValues._SourceItemModified)-Format "yyyy-MM-ddTHH:mm:ssZ")")
                                }

                                $add = $del = $upd = 0;
                                #Lets do the ven diagram thingy to see if we can find the ones to delete / add / (we will do update in a min with another diff calculation) 
                                $DiffIDs = Compare-Object -ReferenceObject $ItemGUIDs  -DifferenceObject $FollowerItemGUIDs 
                                if($DiffIDs){
                                    foreach ($D in $DiffIDs) {
                                        switch ($D.SideIndicator) {
                                        "<=" {
                                                ## ITS AN ADD
                                                #so lets fish it out of the Payload.ItemData
                                                $ER = $payLoad.ItemData.Values | Where-Object { $_._SourceItemGUID -eq $D.InputObject } 
                                                ##added this is to alllow the save 
                                                $ER.Remove("Modified")
                                                #do the add 
                                                $a = Add-PnPListItem -connection $UpdateConnection -List $p.LibPath -Values $ER -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                                                if ($Errvar[0].Count -gt 0) {
                                                    logActivity -Indent 4 -Type "5 Warning" -Message "Problem Adding item $($ER._SourceItemGUID)"
                                                }
                                                else {
                                                    logActivity -Indent 4 -Type "4 Action" -Message "Added item [$($a.Fieldvalues.Title)] matching item $($ER.fields.SourceItemGUID)"
                                                    $add++
                                                }
                                            }
                                
                                        "=>" {
                                                ## ITS A DELETE
                                                #logActivity -Indent 4 -Type "4 Action" -Message "Item $($D.InputObject) should be deleted"
                                                #this is two step process first we get it using the source ITEM Id to find the actual ID then we can delete by The actual ID
                                                $DR = $FollowerItems | Where-Object { $_.FieldValues._SourceItemGUID -eq $D.InputObject } 
                                                if($DR.Count -eq  1){ 
                                                    if ($null -ne $DR) {
                                                        ##only do this if we found one else with the "-force" it will trash the ENTIRE  list contents !!!
                                                        $a = Remove-PnPListItem -Connection $UpdateConnection -List $p.LibPath -Identity $DR.Id -Force -Recycle -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                                                        if ($Errvar[0].Count -gt 0) {
                                                            logActivity -Indent 4 -Type "5 Warning" -Message "Problem Deleting item No: $($D.InputObject)"
                                                        } else {
                                                            logActivity -Indent 4 -Type "4 Action" -Message "Deleted item No: $($DR.Id)"
                                                            $del++
                                                        }
                                                    }
                                                } else {
                                                    logActivity -Indent 4 -Type "2 Info" -Message  "$($DR.Count) Items found" -ForegroundColor Red
                                                    for($UI=0; $UI -lt $DR.Count; $UI++){ 
                                                        $a = Remove-PnPListItem -Connection $UpdateConnection -List $p.LibPath -Identity $DR[$UI].Id -Force -Recycle -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                                                        logActivity -Indent 4 -Type "4 Action" -Message "Deleted item No: $($DR[$UI].Id)" 
                                                    }
                                                }
                                            }
                                        }    
                                    }
                                } else {
                                    logActivity -Indent 2 -Type "2 Info" -Message "Nothing to be added or deleted"
                                }

                                # lets deal with the updates 
                    
                                $DiffDates = Compare-Object -ReferenceObject $ItemGUIDDates  -DifferenceObject $FollowerItemGUIDDates -IncludeEqual
                                if($DiffDates){
                                    foreach($DD in $DiffDates){
                                        if($DD.SideIndicator -eq "<="){
                                            #oh an update so we need to get the data from the source and the ID from the matching one in the follower
                                            $ER = $payLoad.ItemData.Values   | Where-Object { $_._SourceItemGUID -eq $DD.InputObject.split(" ")[0] }
                                            $UpdateableItem = $FollowerItems | Where-Object { $_.FieldValues._SourceItemGUID -eq $DD.InputObject.split(" ")[0] } 
                                            if($UpdateableItem.Count -eq 1) { # TEST an EDGE CASE where a GUID appears more than once 
                                                $a = Set-PnPListItem -connection $UpdateConnection -List $p.LibPath -identity $UpdateableItem.Id  -Values $ER -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                                                if ($Errvar[0].Count -gt 0) {
                                                    logActivity -Indent 4 -Type "5 Warning" -Message "Problem Updating item No: $($DD.InputObject.split(" ")[0])"
                                                } else {
                                                    logActivity -Indent 4 -Type "4 Action" -Message "Updated : $($UpdateableItem.FieldValues.Title)"
                                                    $upd++
                                                }
                                            } else {
                                                #Lets trash them all let it get corrected next time it runs 
                                                for($UI=0; $UI -lt $UpdateableItem.Count; $UI++){ 
                                                    $a = Remove-PnPListItem -Connection $UpdateConnection -List $p.LibPath -Identity $UpdateableItem[$UI].Id -Force -Recycle -ErrorAction SilentlyContinue -ErrorVariable ErrVar
                                                    logActivity -Indent 4 -Type "4 Action" -Message "Deleted duplicate item No: $($UpdateableItem[$UI].Id)" 
                                                }
                                            }
                                        }
                                    }
                                    LogActivity -Indent 3 -Type "3 Info" -Message "We had $tot source items, and created $add new records, $upd Updates, and $del deletes." 
                                } else {
                                   logActivity -Indent 2 -Type "2 Info" -Message "Nothing to Updated"
                                }
                            } else {
                                logActivity -Indent 4 -Type "4 Warning" -Message  "We couldnt find the follower list $f"
                            }
                            logActivity -Indent 3 -Type "2 Info" -Message  "Finished with the follower $f"
                        }
                    }
                }
            }
             
        }  else {
           logActivity -Indent 0 -Type "6 Error" -Message  "cant connect to  - $($SourceSiteUrl) Ending..." #-LogAction "Write"
        }
        $sourceConnection = $null
        #log that it ran - may knock this log level to writequiet later if its not needed but for now its proof for Jen and Shin
        logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished for $SourceSiteUrl " -logAction "Write"
    }
}

#======================================================================================================================================
# TASK END HERE - REGISTER THE PING AND CLOSE OUT THE LOG
#======================================================================================================================================
$p = basicPing

#logActivity -Indent 0 -Type "2 Info" -Message "JOB Finished" -logAction "Write" #WriteQuiet - only write a record IF the max error exceeds the limit  
