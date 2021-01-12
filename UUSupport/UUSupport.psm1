$FormatEnumerationLimit = -1


$MODULEDATE = get-childitem -path "C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\uusupport\uusupport.psm1"
show-line
Write-host  "                UU Support Module loaded`n`n" -ForegroundColor green
Write-host  "         Last changed: "$MODULEDATE.LastWriteTime -ForegroundColor GREEN
show-line


Function get-help3
{
    <#
    #... Displays the Functions in quuSupport.
    Syntax   Get-help1

    This will generate the report for all of the batchname starting with Batch14*
#>

    param
    (

    )

    BEGIN
    {
        Write-Host $version -ForegroundColor Green
    }

    PROCESS
    {
        #... Display Functions with the comments
        $DDSPLAY = ""
        #Reads in the Current Powershell script file
        $DDSPLAY = get-content $PSCommandPath

        foreach ($line in $DDSPLAY)
        {
            if ($line.StartsWith('Function', "CurrentCultureIgnoreCase") -or $line.Trim().Startswith('#...', "CurrentCultureIgnoreCase"))
            {
                $1 = $line
            
            
                if ($1.Trim().StartsWith('#...', "CurrentCultureIgnoreCase"))
                {

                    #Removes the first 4 char's Eg "#... "
                    $linedes = $line.trim().substring(4)
                    Write-Host $linedes -f Gray -NoNewline
                }
                Elseif (!($1.Trim().Startswith('#...')))
                {
                    Write-host ''
                }
                if ($1.Trim().StartsWith('Function', "CurrentCultureIgnoreCase"))
                {
                    $linelong = $line + "                                               "

                    #makes the Line length to be 50 so the comments all line up. Fills it up with a space.
                    $line = $linelong.substring(0, 50)
                    Write-host "  $line" -f green -NoNewline
                }
            }
        }

    }
    END
    {
        Write-Output  ""
    }
}


function Compare-uuUPNandPrimarySMTP
{
    <#
 #... Compares Primary SMTP address and the UPN for mailbox
# # .SYNOPSIS
    Compares Primary SMTP address and the UPN for mailbox
.DESCRIPTION
    Compares Primary SMTP address and the UPN for mailbox.  This will compare either a single mailbox, Array or it will just search all mailboxes.
.PARAMETER one
    Mailboxes

.EXAMPLE
    C:\PS>Compare-UPNandPrimarySMTP -mailboxes lawrence.mckay@urbanutilities.com.au
    C:\PS>Compare-UPNandPrimarySMTP
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    the following info is displayed
                        SamAccount
                        Displayname
                        UserPrincipalname
                        PrimarySmtpAddress
                        UPNPrimarySMTPMatch
.NOTES

    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty

    Date:    15 May 2019

     ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1          31 Jan  2020         Lawrence       Initial Coding

#>
    
    [CmdletBinding()]
    param(                                                                                                     
        $MAilboxes 
    )               
         
    begin
    {
        if ($mailboxes -eq $null)
        {
            
            write-host "getting all the Mailbox info.  This may take a few min" -ForegroundColor Cyan
            $mailboxes = Get-mailbox -resultsize unlimited | Select-Object alias, displayname, userprincipalname, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, IsMailboxEnabled, HiddenFromAddressListsEnabled
        }
        else
        {
            $mailboxes = get-mailbox $mailboxes -resultsize unlimited | Select-Object alias, displayname, userprincipalname, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, IsMailboxEnabled, HiddenFromAddressListsEnabled
        }     
        $i = 0 
    }               
    process 
    { 
        try
        {
            foreach ($mailbox in $mailboxes)
            {
                if ($mailbox.UserPrincipalname -eq $mailbox.PrimarySmtpAddress)
                {
                    #        write-host $online.userprincipalName ' is the same both on Prem and Online' -ForegroundColor green 
                    $Matched = 'MAtched' 
                } 
                else  
                {
                    $MAtched = 'Different'  
                } 
               
                $upn = $mailbox.userprincipalname.toString()
                $userinfo = Get-ADUser -Filter 'UserPrincipalName -eq $upn' -Properties * -ErrorAction SilentlyContinue  
               
                
                <# $MBSize = Get-MailboxStatistics $mailbox.PrimarySmtpAddress.tostring() -ErrorAction silentlycontinue | Select-Object -ExpandProperty Totalitemsize
                $MBSizeArchive = Get-MailboxStatistics $mailbox.PrimarySmtpAddress.tostring() -archive -ErrorAction silentlycontinue | Select-Object -ExpandProperty Totalitemsize 
                $upn = $mailbox.userprincipalname.toString()
                $userinfo = Get-ADUser -Filter 'UserPrincipalName -eq $upn' -Properties * -ErrorAction SilentlyContinue  
                
                $mbsazemb = "N/A"
                if ($MBSize.value) { $mbsazemb = $MBSize.value.ToMB() }
               
                $mbsazeArchmb = "N/A" 
                if ($MBSizeArchive.value)
                { $mbsazeArchmb = $MBSizeArchive.value.ToMB() }

#>


                [PSCustomObject] @{
                    SamAccount           = $mailbox.alias 
                    Displayname          = $mailbox.displayname
                    UserPrincipalname    = $mailbox.userprincipalname.toString()
                    PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress.tostring()
                    UPNPrimarySMTPMatch  = $matched
                    RecipientTypeDetails = $mailbox.RecipientTypeDetails
                    RecipientType        = $mailbox.RecipientType
                    IsMailboxEnabled     = $mailbox.IsMailboxEnabled
                    <#
                    ActiveDirectoryAccountEnabled = $userinfo.Enabled
                    PasswordLastChanged           = $userinfo.PasswordLastSet
                    HiddenFromAddressListsEnabled = $mailbox.HiddenFromAddressListsEnabled
                    Title                         = $userinfo.Title
                    Company                       = $userinfo.company
                    Department                    = $userinfo.Department                                                   
                    Office                        = $userinfo.Office                                                       
                    "Users Manager"               = $userinfo.Manager                                                      
                #  PrimaryMailboxSize            = $mbsazemb                                                              
                #  ArchiveMailboxSize            = $mbsazeArchmb                                                   
                #>
                } 

                if ($mailboxes.count -gt 1)
                {
                    $paramWriteProgress = @{
                        Activity        = 'Exporting Mailbox Informaiton'
                        Status          = "Processing [$i] of [$($mailboxes.Count)] users"
                        PercentComplete = (($i / $mailboxes.Count) * 100)
                        #CurrentOperation = "Completed : [$mailbox.displayname.tostring()]"
                    }
                    Write-Progress @paramWriteProgress
                }
                $i++
                
                
            }
        }
        catch
        {
            $errormsg = "ERROR : $($_.Exception.Message)"
            return $errormsg 
        
    
        }
    }
    
    end
    {
    
    }
}


function  get-uuDistrubutionGroupMembers
{
    <#
 #... get all Distrubution Group Members
.SYNOPSIS
    get all Distrubution Group Members useing the Pram value of name
.DESCRIPTION
    Long description
.PARAMETER Groups
   $groups  

.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 July 2019         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The Output Folder')]
        $Groups 

    )
    
    
    begin 
    {
        $i = 1
    }
    
    process 
    {
            
        try 
        {
            foreach ($group in $Groups)
            {


                if ($Groups.count -gt 1)
                {
                    $paramWriteProgress = @{
                        Activity        = 'ProcessingGroups'
                        Status          = "Processing [$i] of [$($groups.Count)] roles."
                        PercentComplete = (($i / $groups.Count) * 100)
                    }
                    Write-Progress @paramWriteProgress
                }
                $i++

                $DistrubutionGroupMembership = Get-DistributionGroupMember $group -resultsize unlimited -erroraction SilentlyContinue 

                foreach ($user in $DistrubutionGroupMembership)
                {
    
                 
                    [PSCustomObject] @{ 
                        Distribution_GroupName   = $Group.name
                        Group_PrimarySmtpAddress = $Group.PrimarySmtpAddress
                        GroupType                = $Group.RecipientTypeDetails
                        UserDisplayname          = $user.DisplayName
                        UserSamAccount           = $user.alias
                        UserRecipientType        = $user.RecipientTypeDetails
                        UserPrimarySMTPAddress   = $user.PrimarySMTPAddress
                        
                    }
                }
            }


        }
        catch 
        {
            $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
            Write-Host $Errormsg -ForegroundColor magenta
        }
    }
    
    end
    {
            
    }
}

function get-uuMailboxFolderCountOnPrem
{
    <#
#... Gets all the Folder counts for all mailbox from Pram.
	.SYNOPSIS
		A brief description of the get-quuMailboxFolderCountOnPrem function.

	.DESCRIPTION
		A detailed description of the get-quuMailboxFolderCountOnPrem function.

	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.

	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.

	.EXAMPLE
		PS C:\> get-quuMailboxFolderCountOnPrem

	.NOTES
		Additional information about the function.
#>
    [CmdletBinding()]
    param
    (
        [String[]]$UserPrincipalName,
        [Switch]$ShowProgress
    )

    begin
    {
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
        $i = 1
        $DomainController = Get-ADDomainController | Select-Object -ExpandProperty hostname

    }
    process
    {
        foreach ($UPN in $UserPrincipalName)
        {
            try
            {
                $Recipient = Get-Recipient $upn  -DomainController  $DomainController  -ErrorAction Stop
                $mbxStats = Get-MailboxFolderStatistics $UPN -DomainController $DomainController -ErrorAction Stop
                $adinfo = get-aduser -Filter { UserPrincipalName -eq $upn } | select -ExpandProperty UserPrincipalName
                $prop = [ordered]@{
                    Samaccountname       = $Recipient.samaccountname
                    Displayname          = $Recipient.DisplayName
                    UserPrincipalName    = $adinfo
                    PrimarySmtpAddress   = $Recipient.PrimarySmtpAddress
                    Title                = $Recipient.title
                    Department           = $recipient.Department
                    Office               = $Recipient.office
                    Company              = $Recipient.company
                    RecipientType        = $Recipient.RecipientType
                    RecipientTypeDetails = $Recipient.RecipientTypedetails
                    
                    Mailbox              = 'OnPremise'
                    TotalFolders         = $mbxStats.Count
                    Details              = 'None'

                }
                #start-sleep 1
            }
            catch
            {
                $prop = [ordered]@{
                    UserPrincipalName = $UPN
                    Mailbox           = 'OnPremise'
                    TotalFolders      = 'ERROR'
                    Details           = "$($_.Exception.Message)"
                }
            }
            finally
            {
                $obj = New-Object -TypeName psobject -Property $prop
                Write-Output $obj

                if ($ShowProgress)
                {
                    if ($UserPrincipalName.count -gt 1)
                    {
                        $paramWriteProgress = @{
                            Activity         = 'Counting Folders for Mailboxes Onpremise'
                            Status           = "Processing [$i] of [$($UserPrincipalName.Count)] users"
                            PercentComplete  = (($i / $UserPrincipalName.Count) * 100)
                            CurrentOperation = "Completed : [$UPN]"
                        }
                        Write-Progress @paramWriteProgress
                    }
                }
                $i++
            }

        }
    }
    end
    {
        Write-Progress -Activity 'Counting Folders for Mailboxes' -Completed
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
    }
}


function get-uuSharedMBPermissions
{ 
    <#
 #... get Shared MB Permissions
.SYNOPSIS
    Get Shared MAilbox Permissions for mailboxes
.DESCRIPTION
    Get shared mailbox permissions
    eg   Full Read

    It is excluding the following users..  
            "*Managed Availability Servers"
            "*Delegated Setup"
            "*Domain Admins"
            "*Enterprise Admins"
            "*Exchange Servers"
            "*Exchange Trusted Subsystem"
            "*Organization Management"
            "*Public Folder Management"
    
.PARAMETER Bulkuser
    USer / Users in an array

.EXAMPLE
    C:\PS>get-quuSharedMBPermissions -SharedMB <Displayname>
    C:\PS>get-quuSharedMBPermissions
    


.NOTES

Needs a connection to Exchange, AD and MSOL Service

    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           13 Feb 2020         Lawrence       Initial Coding

#>
    
    [CmdletBinding()]
    param
    (
        $SharedMB, 
        [Switch]$ShowProgress
    )

    begin 
    {
        
        if ($SharedMB -eq $null)
        {
            
            write-host "getting all the Mailbox info.  This may take a few min" -ForegroundColor Cyan
            $global:SharedMB = Get-mailbox  -ResultSize unlimited | where { $_.RecipientTypeDetails -match "shared" } | Select-Object alias, displayname, userprincipalname, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, IsMailboxEnabled, HiddenFromAddressListsEnabled
        }
      
        $i = 1 
    }
    
    process 
    {
        try 
        {
            foreach ($name in $global:SharedMB)
            {
             
                if ($ShowProgress)
                {
                    #sHOW PROGRESS bAR
                    $paramWriteProgress = @{
                        Activity        = 'Processing the Shared MAilboxes  '
                        Status          = "Processing  [$i] of [$($SharedMB.Count)] Mailboxes"
                        PercentComplete = (($i / $SharedMB.Count) * 100)

                    }

                    Write-Progress @paramWriteProgress
                }
                $i++
                #Write-host "MB Name" $name.alias
                $MBPerms = Get-Mailboxpermission -identity $name.alias | Where-Object { ($_.User -notlike "*ather.malik") -and ($_.User -notlike "*ari.perdis") -and ($_.User -like "urbanutilities*") -and ($_.user -notlike "*Managed Availability Servers") -and ($_.user -notlike "*Delegated Setup") -and ($_.user -notlike "*Domain Admins") -and ($_.user -notlike "*Enterprise Admins") -and ($_.user -notlike "*Exchange Servers") -and ($_.user -notlike "*Exchange Trusted Subsystem") -and ($_.user -notlike "*Organization Management") -and ($_.user -notlike "*Public Folder Management") } 
                #write-host "1"
                $MBInfo = Get-Recipient  $name.primarySMTPAddress.tostring()  -ErrorAction silentlycontinue
                #write-host "2"
                [string]$UP = $name.userprincipalname
                #write-host "3"
                $MBInfoAD = get-aduser  -filter { UserPrincipalName -eq $up }
                #write-host "4"
                foreach ($line in $mbperms)
                {
                   
                    #    write-host ".1"               
                    $userinfo = $null
                    if ($line.user.RawIdentity -like "urbanutilities*")
                    {
                        #       write-host ".2"    
                        $e3license = $null                        
                        [string]$permis = $line.AccessRights
                        [string]$SAM = $line.user.RawIdentity.split("\")[-1]
                        $userinfo = get-aduser -Filter "samaccountname -eq '$sam'" -Properties *
                        
                        [string]$upn = $userinfo.userprincipalname
                        if ($upn -match "@")
                        {
                            
                            $E3Lic = get-msoluser -UserPrincipalName $userinfo.UserPrincipalName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Licenses | select-object -ExpandProperty AccountSkuId  
                            if ($E3Lic -match 'queenslandurbanutilities:ENTERPRISEPACK' ) { $e3license = 'E3 License' }
                            $Recipienttype = Get-Recipient $SAM -ErrorAction silentlycontinue
                            #            write-host ".3"
                        }
                    
                        $userinfoSamaccountname = $userinfo.Samaccountname
                        $userinfoEnabled = $userinfo.Enabled
                        $userinfoDisplayname = $Recipienttype.Displayname
                        $userinfocompany = $Recipienttype.company
                        $userinfoDepartment = $userinfo.Department
                        $userinfoOffice = $Recipienttype.Office
                        $RecipientTyp = $Recipienttype.RecipientTypeDetails
                        
                        #setting Recipient Type to not found if is Empty
                        #.Cleanup this bit
                        #       write-host ".4"
                        if (!($Recipienttype))
                        {
                            $RecipientTyp = "MB not found"
                            #           write-host ".5"
                        }
                        #        write-host "."

                        if (!($userinfo ))
                        {
                            #          write-host ".6"
                            $GRP = get-adgroup $sam -ErrorAction silentlycontinue  
                            #           write-host ".7"
                            $GRPTYPE = $g.GroupCategory#($g.GroupScope).ToString() + " - " + ($g.GroupCategory).ToString()
                            #            write-host ".8"
    
                            $userinfoSamaccountname = $grp.Samaccountname
                            $userinfoEnabled = $userinfo.Enabled
                            $userinfoDisplayname = $grp.Displayname
                            $userinfocompany = $userinfo.company
                            $userinfoDepartment = $userinfo.Department
                            $userinfoOffice = $userinfo.Office
                            $RecipientTyp = $GRPTYPE

                        }
                        #        write-host ".9"

                        [PSCustomObject] @{ 
                            MailboxSamAccountName      = $MBInfo.samaccountname
                            Displayname                = $mbinfo.displayname

                            PrimarySmtpAddress         = $mbinfo.primarySMTPAddress
                            UserWithPermissionsSam     = $userinfoSamaccountname
                            UserWithPermDisplayname    = $userinfoDisplayname
                            UserWithPermAccountEnabled = $userinfoEnabled
                            UserwithPermE3Licensed     = $e3license
                            UserWithPermCompany        = $userinfocompany
                            UserWithPermDepartment     = $userinfoDepartment
                            UserWithPermOffice         = $userinfoOffice
                            Permissions                = $permis   
                            RecipientTypeDetails       = $RecipientTyp
                            
                        }
                   
                    }
                }
            }
                
 

        }

        Catch 
        {
            $errormsg = "ERROR : $($_.Exception.Message)"
            return $errormsg 
        }
    }
}

function get-uuUserMBSharePermissions
{ 
    <#
 #... get Shared MB Permissions
.SYNOPSIS
    Get Shared MAilbox Permissions for mailboxes
.DESCRIPTION
    Get shared mailbox permissions
    eg   Full Read

    It is excluding the following users..  
            "*Managed Availability Servers"
            "*Delegated Setup"
            "*Domain Admins"
            "*Enterprise Admins"
            "*Exchange Servers"
            "*Exchange Trusted Subsystem"
            "*Organization Management"
            "*Public Folder Management"
    
.PARAMETER Bulkuser
    USer / Users in an array

.EXAMPLE
    C:\PS>get-quuSharedMBPermissions -SharedMB <Displayname>
    C:\PS>get-quuSharedMBPermissions
    


.NOTES

Needs a connection to Exchange, AD and MSOL Service

    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           13 Feb 2020         Lawrence       Initial Coding

#>
    
    [CmdletBinding()]
    param
    (
        $SharedMB, 
        [Switch]$ShowProgress
    )

    begin 
    {
        
        if ($SharedMB -eq $null)
        {
            
            write-host "getting all the Mailbox info.  This may take a few min" -ForegroundColor Cyan
            $SharedMB = Get-mailbox  -ResultSize unlimited | Select-Object alias, displayname, userprincipalname, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, IsMailboxEnabled, HiddenFromAddressListsEnabled
        }
      
        $i = 1 
    }
    
    process 
    {
        try 
        {
            foreach ($name in $SharedMB)
            {
             
                if ($ShowProgress)
                {
                    #sHOW PROGRESS bAR
                    $paramWriteProgress = @{
                        Activity        = 'Processing the Shared MAilboxes  '
                        Status          = "Processing  [$i] of [$($SharedMB.Count)] Mailboxes"
                        PercentComplete = (($i / $SharedMB.Count) * 100)

                    }

                    Write-Progress @paramWriteProgress
                }
                $i++
                Write-host "MB Name" $name.alias
                $MBPerms = Get-Mailboxpermission -identity $name.alias | Where-Object { ($_.User -notlike "*ather.malik") -and ($_.User -notlike "*ari.perdis") -and ($_.User -like "urbanutilities*") -and ($_.user -notlike "*Managed Availability Servers") -and ($_.user -notlike "*Delegated Setup") -and ($_.user -notlike "*Domain Admins") -and ($_.user -notlike "*Enterprise Admins") -and ($_.user -notlike "*Exchange Servers") -and ($_.user -notlike "*Exchange Trusted Subsystem") -and ($_.user -notlike "*Organization Management") -and ($_.user -notlike "*Public Folder Management") } 
                write-host "1"
                $MBInfo = Get-Recipient  $name.primarySMTPAddress.tostring()  -ErrorAction silentlycontinue
                write-host "2"
                [string]$UP = $name.userprincipalname
                write-host "3"
                $MBInfoAD = get-aduser  -filter { UserPrincipalName -eq $up }
                write-host "4"
                foreach ($line in $mbperms)
                {
                   
                    write-host ".1"               
                    $userinfo = $null
                    if ($line.user.RawIdentity -like "urbanutilities*")
                    {
                        write-host ".2"    
                        $e3license = $null                        
                        [string]$permis = $line.AccessRights
                        [string]$SAM = $line.user.RawIdentity.split("\")[-1]
                        $userinfo = get-aduser -Filter "samaccountname -eq '$sam'" -Properties *
                        
                        [string]$upn = $userinfo.userprincipalname
                        if ($upn -match "@")
                        {
                            
                            $E3Lic = get-msoluser -UserPrincipalName $userinfo.UserPrincipalName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Licenses | select-object -ExpandProperty AccountSkuId  
                            if ($E3Lic -match 'queenslandurbanutilities:ENTERPRISEPACK' ) { $e3license = 'E3 License' }
                            $Recipienttype = Get-Recipient $SAM -ErrorAction silentlycontinue
                            write-host ".3"
                        }
                    
                        $userinfoSamaccountname = $userinfo.Samaccountname
                        $userinfoEnabled = $userinfo.Enabled
                        $userinfoDisplayname = $Recipienttype.Displayname
                        $userinfocompany = $Recipienttype.company
                        $userinfoDepartment = $userinfo.Department
                        $userinfoOffice = $Recipienttype.Office
                        $RecipientTyp = $Recipienttype.RecipientTypeDetails
                        
                        #setting Recipient Type to not found if is Empty
                        #.Cleanup this bit
                        write-host ".4"
                        if (!($Recipienttype))
                        {
                            $RecipientTyp = "MB not found"
                            write-host ".5"
                        }
                        write-host "."

                        if (!($userinfo ))
                        {
                            write-host ".6"
                            $GRP = get-adgroup $sam -ErrorAction silentlycontinue  
                            write-host ".7"
                            $GRPTYPE = $g.GroupCategory#($g.GroupScope).ToString() + " - " + ($g.GroupCategory).ToString()
                            write-host ".8"
    
                            $userinfoSamaccountname = $grp.Samaccountname
                            $userinfoEnabled = $userinfo.Enabled
                            $userinfoDisplayname = $grp.Displayname
                            $userinfocompany = $userinfo.company
                            $userinfoDepartment = $userinfo.Department
                            $userinfoOffice = $userinfo.Office
                            $RecipientTyp = $GRPTYPE

                        }
                        write-host ".9"

                        [PSCustomObject] @{ 
                            MailboxSamAccountName      = $MBInfo.samaccountname
                            Displayname                = $mbinfo.displayname

                            PrimarySmtpAddress         = $mbinfo.primarySMTPAddress
                            UserWithPermissionsSam     = $userinfoSamaccountname
                            UserWithPermDisplayname    = $userinfoDisplayname
                            UserWithPermAccountEnabled = $userinfoEnabled
                            UserwithPermE3Licensed     = $e3license
                            UserWithPermCompany        = $userinfocompany
                            UserWithPermDepartment     = $userinfoDepartment
                            UserWithPermOffice         = $userinfoOffice
                            Permissions                = $permis   
                            RecipientTypeDetails       = $RecipientTyp
                            
                        }
                   
                    }
                }
            }
                
 

        }

        Catch 
        {
            $errormsg = "ERROR : $($_.Exception.Message)"
            return $errormsg 
        }
    }
}


function get-uuMailboxFolderByItemCount
{
    <#
#... Gets all the Folder Item counts for user mailbox
	.SYNOPSIS
		Gets the number of Email items in the Folders and Sorts them to be low to high.

	.DESCRIPTION
        Gets the number of Email items in the Folders and Sorts them to be low to high.
        This retrieves all the items including hidden items so there may appear to be 
        some items in the folders but they don't exist to the USers.

        Users ImportExcel PS module
        To install ---> Install-Module ImportExcel

        

	.PARAMETER UserPrincipalName
        This is the user UPN or SAM   Will also work with an array
     
     .PARAMETER OutPutPath 
     the path to write the output files to. (CSV and XLSX files.)  

	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.

	.EXAMPLE
        PS C:\> get-quuMailboxFolderByItemCount -UserPrincipalName <UPN or SAM> -OutPutPath <c:\tmp>
        PS C:\> get-quuMailboxFolderByItemCount -UserPrincipalName <UPN or SAM> -showProgress -OutPutPath <c:\tmp>
        PS C:\> get-quuMailboxFolderByItemCount -UserPrincipalName $Bulkusers       
        PS C:\> get-quuMailboxFolderByItemCount -UserPrincipalName $Bulkusers |out-file xyz.txt       

	.NOTES
        This function requires Exchange tools to run.

        Lawrence McKay
		Lawrence@mckayit.com
        McKayIT Solutions Pty 
        
        Date:    15 May 2018
        
        ******* Update Version number below when a change is done.*******
        
        History
        Version         Date                Name           Detail
        ---------------------------------------------------------------------------------------
        0.0.1           09 MArch 2020        Lawrence       Initial Coding
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The UPN')]
        $UserPrincipalName,
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The Path where files are to be created')]
        $OutPutPath,
        
        [Switch]$ShowProgress
    )

    begin
    {
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
        $i = 1
        #$DomainController = Get-ADDomainController | Select-Object -ExpandProperty hostname

    }
    process
    {
    
        foreach ($UPN in $UserPrincipalName)
        {
            try
            {
                #progress
                if ($ShowProgress)
                {
                    if ($UserPrincipalName.count -gt 1)
                    {
                        $paramWriteProgress = @{
                            Activity         = 'Counting Folders for Mailboxes Onpremise'
                            Status           = "Processing [$i] of [$($UserPrincipalName.Count)] users"
                            PercentComplete  = (($i / $UserPrincipalName.Count) * 100)
                            CurrentOperation = "Completed : [$UPN]"
                        }
                        Write-Progress @paramWriteProgress
                    }
                }
                $i++


                #$Recipient = Get-Recipient $upn  -DomainController  $DomainController  -ErrorAction Stop
                $Recipient = Get-Recipient $upn   -ErrorAction Stop
                Write-host ' '
                Write-host ' '
                Write-host ' Here is the Folder item count for Mailbox: ' $Recipient.DisplayName -ForegroundColor Green
                $folderstats = Get-MailboxFolderStatistics -Identity $upn | Select-Object folderpath, ItemsInFolder | Sort-Object ItemsInFolder
                
                #changes to Folder
                cd $OutPutPath
                # Sets the output files 
                $filenamecsv = $OutPutPath + "\" + $upn.Split("@")[0] + ".csv"
                $filename = $OutPutPath + "\" + $upn.Split("@")[0] + ".xlsx"
                
                $folderstats | Export-csv $filenamecsv -NoTypeInformation
                
                import-csv $filenamecsv | Export-Excel $filename -AutoSize -TableName foldercount -WorksheetName Foldercount



            
            }
            catch
            {
                $prop = [ordered]@{
                    UserPrincipalName = $UPN
                    Mailbox           = 'OnPremise'
                    TotalFolders      = 'ERROR'
                    Details           = "$($_.Exception.Message)"
                }
                $obj = New-Object -TypeName psobject -Property $prop
                Write-Output $obj
            }
            finally
            {
               

               
            }

        }
    }
    end
    {
        
        $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
    }
}


function get-Allmailbox_Primary_and_ProxyAddress_with_SAM
{
    <#
 #... Get All Mailbox Primary and Proxy Addresses
.SYNOPSIS
    Get All Mailbox sufix's
.DESCRIPTION
    gets input of users UPN and then returnes all Email Alias Sufix  Eg after the @ symbol

.PARAMETER UPNs
    UPN or array of users UPN

.PARAMETER showprogress
    Switch to sho progress

    
.EXAMPLE
    C:\PS>get-Allmailbox_Primary_and_ProxyAddress_with_SAM -UPNs <UPN>
    C:\PS>get-Allmailbox_Primary_and_ProxyAddress_with_SAM -UPNs <UPN> -showprogress
    C:\PS>get-Allmailbox_Primary_and_ProxyAddress_with_SAM -UPNs <UPN> -showprogress |export-csv <file>.csv
    
    

.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           12 March  2019        Lawrence     Initial Coding

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter Users or Array')]
        $UPNs,
        [Switch]$ShowProgress
    
    )

    Begin
    {
        $i = 0
    }
    process 
    {

        foreach ($user in $UPNs)
        {
            try
            {
                if ($ShowProgress)
                {
                    #sHOW PROGRESS bAR
                    $paramWriteProgress = @{
                        Activity        = 'Processing  MAilboxes  '
                        Status          = "Processing  [$i] of [$($UPNs.Count)] Mailboxes"
                        PercentComplete = (($i / $UPNs.Count) * 100)

                    }

                    Write-Progress @paramWriteProgress
                    $i++
                }
                

                $Usersmtps = get-mailbox $user  -erroraction silentlycontinue  

                foreach ($userSmtp in $usersmtps)
                {

                    foreach ($caseSMTP in $userSmtp.emailaddresses)
                    {
                        $Proxyaddress = $null
                        if ($casesmtp.PrefixString -clike 'smtp')
                        {
                            $Proxyaddress = $casesmtp.smtpaddress
                            [PSCustomObject] @{ 
                                Samaccountname = $userSmtp.SamAccountName
                                PrimarySMTP    = $userSmtp.PrimarySmtpAddress
                                Proxyaddress   = $Proxyaddress
                            }
                        }
                    }
                }
            }
            catch
            {
                $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
                Write-Host $Errormsg -ForegroundColor magenta
            }
        }
    }
    end
    {

    }
}

function get-passwordexpireDetails
{ 
    <#
 #... get-passwordexpireDetails
.SYNOPSIS
    get-passwordexpireDetails
.DESCRIPTION
    get-passwordexpireDetails
.PARAMETER Bulkuser
    USer / Users in an array

.EXAMPLE
    C:\PS>get-passwordexpireDetails -Users $users
    C:\PS>get-passwordexpireDetails -Users lawrence.mckay@urbanutilities.com.au
    C:\PS>get-passwordexpireDetails -Users 102838


.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           24 Mar 2020         Lawrence       Initial Coding

#>
    
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $Users

    )


    begin 
    {
        $mbtContext = @{
            "1"       = "User Mailbox"
            "2"       = "Linked Mailbox" 
            "4"       = "Shared Mailbox"
            "8"       = "Legacy Mailbox"  
            "16"      = "Room Mailbox"
            "32"      = "Equipment Mailbox"  
            "64"      = "Mail Contact"
            "128"     = "Mail-enabled User"  
            "256"     = "Mail-enabled Universal Distribution Group"
            "512"     = "Mail-enabled non-Universal Distribution Group" 
            "1024"    = "Mail-enabled Universal Security Group"   
            "2048"    = "Dynamic Distribution Group"   
            "4096"    = "Mail-enabled Public Folder"  
            "8192"    = "System Attendant Mailbox"   
            "16384"   = "Mailbox Database Mailbox"  
            "32768"   = "Across-Forest Mail Contact" 
            "65536"   = "User" 
            "131072"  = "Contact"  
            "262144"  = "Universal Distribution Group"  
            "524288"  = "Universal Security Group"  
            "1048576" = "Non-Universal Group"   
            "2097152" = "Disabled User"    
            "4194304" = "Microsoft Exchange"  
        }
             
    }
    
    process 
    {

        try 
        {

            foreach ($auser in $users)
            {
                $user = get-aduser -Properties * $auser
                #$typ = @()
                #$upn = @()
                $numproxyaddres = 0
                clear-variable -name ('emailaddr' , 'upn', 'typ', 'ProxyEmailaddress', 'numproxyaddres')  -ErrorAction silentlycontinue
                Clear-Variable -Name ('uname', 'Parentou', 'InitialCN')  -ErrorAction silentlycontinue

                if ($user.msExchRecipientTypeDetails -eq $null)
                {
                    $typ = "No mailbox type"
       
        
                }
                else 
                {
                    $typ = $mbtContext["$($user.msExchRecipientTypeDetails)"]
                }

                if ($null -eq $user.UserPrincipalName)
                {
                    $upn = "No UPN"
                    $PasswordExpire = "N/A"
                    
                }
                else 
                {
            
                    $upn = $user.UserPrincipalName
                    $psExpire = $user | select @{Name = "ExpiryDate"; Expression = { [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed") } }
                }


                if ($user.mail -ne "*@*")
                {
                    #$emailaddr = $user.mail.ToString()
                    $emailaddr = "No email address"
                }


                if ($user.mail -match "@")
                {
                    
                    $emailaddr = $user.mail.ToString()
        
                    #$usr = get-aduser $user -Properties CanonicalName, Displayname
                    $cn = $user.CanonicalName
                        
                    $uname = $cn.split('/')[-1]
                    $InitialCN = $cn -split ("/")
                    $ParentOU = $InitialCN[0..$($InitialCN.Count - 2)] -Join "/"
                  


                }
    
                if ($user.proxyaddresses -match '@')
                {
                    Clear-Variable -Name ('ProxyEmailaddress')  -ErrorAction silentlycontinue
                    
                    foreach ($SmtpEmail in $user.proxyAddresses)
                    {
                        
                        if ($SmtpEmail -match 'smtp')
                        {
                            $numproxyaddres++
                            #        write-host $SmtpEmail
                            $ProxyEmailaddress += $SmtpEmail + ' ; '
                        }
                    }
                }
                
                $PasswordExpire = $psExpire.ExpiryDate.tostring("dd-MM-yyyy")

                $prop = [ordered]@{
                    UPN                  = $upn
                    EmailAddress         = $emailaddr.ToString()
                    RecipientTypeDetails = $typ
                    SamAccountName       = $user.SamAccountName
                    CanonicalParentOU    = $ParentOU
                    PasswordExpiry       = $PasswordExpire
                    PasswordNeverExpires = $user.PasswordNeverExpires


                }
                $obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
                Write-Output $obj


                if ($users.count -gt 1)
                {
                    $paramWriteProgress = @{
                        Activity        = 'Getting mailbox details and Proxy addresses '
                        Status          = "Processing $($user.displayname)  [$i] of [$($users.Count)] users"
                        PercentComplete = (($i / $users.Count) * 100)

                    }

                    Write-Progress @paramWriteProgress
                }
                $i++


            }
        }
        Catch 
        {
            $Errormsg = $_.exception.message
            Write-Host $Errormsg -ForegroundColor magenta
        }
              
    }
}


function get-mailboxdump_TrancheAnalysis
{
    <#
 #... get MAilbox info for Tranche Analysys
.SYNOPSIS
    get MAilbox info for Tranche Analysys
.DESCRIPTION
    get MAilbox info for Tranche Analysys
.
.EXAMPLE
    C:\PS>
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 July 2019         Lawrence       Initial Coding

#>
     
    
    begin 
    {

    }
    
    process 
    {
            
        try 
        {
            get-mailbox -resultsize unlimited | select Database, UseDatabaseRetentionDefaults, RetainDeletedItemsUntilBackup, DeliverToMailboxAndForward, LitigationHoldEnabled, SingleItemRecoveryEnabled, RetentionHoldEnabled, RetentionPolicy, AddressBookPolicy, CalendarRepairDisabled, ExchangeGuid, ExchangeUserAccountControl, MessageTrackingReadStatusEnabled, ExternalOofOptions, ForwardingAddress, ForwardingSmtpAddress, RetainDeletedItemsFor, IsMailboxEnabled, OfflineAddressBook, ProhibitSendQuota, ProhibitSendReceiveQuota, RecoverableItemsQuota, RecoverableItemsWarningQuota, DowngradeHighPriorityMessagesEnabled, RecipientLimits, IsResource, IsLinked, IsShared, LinkedMasterAccount, ResourceCapacity, ResourceType, SamAccountName, AntispamBypassEnabled, ServerName, UseDatabaseQuotaDefaults, IssueWarningQuota, RulesQuota, Office, UserPrincipalName, UMEnabled, ThrottlingPolicy, RoleAssignmentPolicy, SharingPolicy, ArchiveDatabase, ArchiveGuid, ArchiveQuota, ArchiveWarningQuota, ArchiveDomain, ArchiveStatus, RemoteRecipientType, DisabledArchiveDatabase, QueryBaseDNRestrictionEnabled, MailboxMoveFlags, MailboxMoveStatus, IsPersonToPersonTextMessagingEnabled, IsMachineToPersonTextMessagingEnabled, CalendarVersionStoreDisabled, SKUAssigned, AuditEnabled, AuditLogAgeLimit, WhenMailboxCreated, UsageLocation, HasPicture, HasSpokenName, Alias, ArbitrationMailbox, OrganizationalUnit, CustomAttribute1, CustomAttribute10, CustomAttribute11, CustomAttribute12, CustomAttribute13, CustomAttribute15, CustomAttribute2, DisplayName, ExternalDirectoryObjectId, HiddenFromAddressListsEnabled, MaxSendSize, MaxReceiveSize, ModerationEnabled, EmailAddressPolicyEnabled, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, RequireSenderAuthenticationEnabled, SimpleDisplayName, SendModerationNotifications, WindowsEmailAddress, MailTip, IsValid, ExchangeVersion, Name, Identity, Guid, ObjectCategory, WhenChanged, WhenCreated, WhenChangedUTC, WhenCreatedUTC, OrganizationId, OriginatingServer


        }
        catch 
        {
            $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
            Write-Host $Errormsg -ForegroundColor magenta
        }
    }
    
    end
    {
            
    }
}



Function new-quubatchNoAutostart2
{
    #... Create new Migration Batch with AutoStart Disabled..
    <#

Syntax
    new-quubatchNoAutostart  -Batchname batch28  -MigrationBatchName 1111  -migrationCSVfile batch28_fix.csv -MRSServers mrs4.quuuper.qld.gov.au

Notes

    The MRS Servers are pre coded to use ('mrs.quuuper.qld.gov.au', 'mrs1.quuuper.qld.gov.au', 'mrs2.quuuper.qld.gov.au', 'mrs3.quuuper.qld.gov.au', 'mrs4.quuuper.qld.gov.au')



    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.



    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd

    Date:    7 Nov 2018


    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           7 Nov 2018       Lawrence       Initial Coding

   #>
    [CmdletBinding()]
    param
    (
        $Batchname,
        [Parameter(Mandatory = $true)]
        $MigrationBatchName,
        [Parameter(Mandatory = $true)]
        $migrationCSVfile
       
    )
    BEGIN
    {
        ## export previous Moves to a Report.
        # Export-QuuMoveReport
        
        Navigate-quuMigrationFolder $batchname
        $adcred = Get-Credential -Credential "l.mckay-admin" 
        $MigrationBatchName = 'lastmoves'
        
        # $UserCredential = Get-Credential -Credential lawrence.mckay@urbanutilities.com.au
        #Start-Sleep 7200
        # $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

        # Import-PSSession $Session -DisableNameChecking -prefix "EXO"
      
    }
    PROCESS
    {
        if (Navigate-quuMigrationFolder $batchname)
        {
            Try
            {
                $location = (Get-Item -Path ".\").FullName + '\' + $migrationCSVfile
                $Batchmember = import-csv $location | select -ExpandProperty emailaddress
                
                $Batchmember | new-exomoverequest  -remote -RemoteHostName "outlook.urbanutilities.com.au" -RemoteCredential $adcred -TargetDeliveryDomain "queenslandurbanutilities.mail.onmicrosoft.com" -SuspendWhenReadyToComplete -baditemlimit 150 -largeitemlimit 100 -BatchName $MigrationBatchName -WarningAction silentlycontinue -acceptlargedataloss
                
            }

            Catch
            {
                # Catches error from Try
                Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
            }
        }
    }
    END
    {

    }


}


function get-batchmovestats
{
    <#
 #... get all moverequest throughput for batch
.SYNOPSIS
    get moverequest Stats for all users in current Batch while Syncing.
.DESCRIPTION
    get moverequest Stats for all users in current Batch while Syncing.
    This will loogup all the move requests currently in a Batch and output the following  every 15 Sec.
                            "Datetime"
                            "Displayname"          
                            "Sam/Alias"            
                            "KBTransferedPerMinute"
                            "PercentageComplete" 

    Can be outputed to CSV if Required.                            
.PARAMETER Batchname
    Specifies the Batch name to monitor
.PARAMETER Delay
    Specifies the Delay between Cycles
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>get-batchmovestats -Batchname <Batchname> -Delay <Timein Sec>
    C:\PS>get-batchmovestats -Batchname Dogfood -Delay 1
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    3 April 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           3rd April 2019      Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The Output Folder')]
        $Batchname,
        $Delay = 15
    )
    
    begin 
    {
        #Used to sets Loop 
        $lpm = 1
    }
    
    process 
    {
        try 
        {
            #get Batch Details    
            try 
            {
                
                $batch = Get-EXOMoverequest -batchname $batchname
            }
            catch 
            {
                    
                Write-Host "$batchname  Batchname not found" -ForegroundColor magenta
            }
                
            do
            {
                #sleep for 30 Sec
                    
                foreach ($user in $batch)
                {
                    $d = Get-Date

                    $stats = Get-EXOMoveRequestStatistics -identity $user.alias | select displayname, alias, BytesTransferredPerMinute, PercentComplete
                    $BTM = [math]::round(($stats.BytesTransferredPerMinute.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1kb), 2)
                    $prop = [ordered]@{
                        "Datetime"              = $d.GetDateTimeFormats()[24]
                        "Displayname"           = $Stats.displayname
                        "Sam/Alias"             = $stats.alias
                        "KBTransferedPerMinute" = $btm
                        "PercentageComplete"    = $stats.PercentComplete
                    

                    }
                    $obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
                    Write-Output $obj
                }

                #delay between Cycles
                Sleep $delay
            }
            until ($lpm -eq 10)

        }
 
        catch 
        {
            $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
            Write-Host $Errormsg -ForegroundColor magenta
        }
    }
    
    end
    {
            
    }
}


function remove-qhdupMailbox
{
    #... Removes the Exchange Attrib for Duplicate mailbox.n.
    <# 
    ********************************************************************************  
    *                                                                              *  
    *              This script IS a Blank Template                                 *
    *                                                                              *  
    ********************************************************************************    
    Note.
        Need to add as exclusion to retention Policy before running.

        
    SYNTAX.
        <remove-qhdupMailbox -mailbox <emailaddress>


      
    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

    

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:    6 Nov 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           6 Nov 2018       Lawrence       Initial Coding
    
   #>

    [CmdletBinding()]
    param 
    (
        $MAilbox
    )
    BEGIN
    {
        
    }
    PROCESS
    {
        
        Try
        {
            Set-EXOMailbox $mailbox -removeDelayHoldApplied
            Set-EXOMailbox $mailbox -RemoveDelayReleaseHoldApplied
            Set-exoMailbox $mailbox -ExcludeFromAllOrgHolds
            get-exoMailbox -Identity $mailbox |fl *hold*
            Pause 
            write-host 'Now Manually Remove E2 and Skype license'
            Write-host "Sleeping for 5 Sec"
            Start-Sleep 5

            Set-EXOUser  $mailbox -PermanentlyClearPreviousMailboxInfo -confirm:$false
            Write-host "Sleeping for 10 Sec"
            Start-Sleep 10
            Write-host  ' The RecipientTypeDetails shoud now be showing as "User"' -ForegroundColor Green
            Write-host  ' If not then wait 10Min and re run Get-EXOUser <UPN>  to check it has changed.' -ForegroundColor Green
            Get-EXOUser $mailbox

        }
            
        Catch
        {
            # Catches error from Try 
            Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
        }
    }
    END
    {
		
    }
}

function test-test
{

    write-host "updated"
}


function Export-QuuMoveReport
{
    <#
 #... Export move Status Report
.SYNOPSIS
    Export move Status Report and saves it as a Excel file
.DESCRIPTION
    Export move Status Report and saves it as a Excel file
    eg "C:\Users\102838\urbanutilities.com.au\Exchange Online - General\Documents\17 MigrationLogs\ " + $timer + "-MoveRequestReportfile.xlsx"
.EXAMPLE
    C:\PS>Export-QuuMoveReport
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    16 April 2029
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           16 April 2020       Lawrence       Initial Coding
    0.0.2           26 May 2020         Lawrence       Updated to include dates for moves
#>
     
    [CmdletBinding()]
    param(

    )
    
    
    begin 
    {
        Write-host  "Exporting moves Report  before excuting new move requests."
        $i = 0
             
    }
    
    process 
    {
            
        try 
        {
            $timer = (Get-Date -Format yyy-MM-dd)
            $filename = "C:\Users\102838\urbanutilities.com.au\Exchange Online - General\Documents\17 MigrationLogs\" + $timer + "-DetailedMoveRequestReportfile.xlsx"
            $filenamecomplete = "C:\Users\102838\urbanutilities.com.au\Exchange Online - General\Documents\17 MigrationLogs\" + $timer + "-CompletedMoveRequestReportfile.xlsx"
            Write-host 'Getting Mailbox move list' -ForegroundColor Cyan
            $id = Get-EXOMoveRequest -ResultSize Unlimited #| select DisplayName, Status, batchname | sort batchname | Export-Excel $filename -AutoSize -TableName MigrationReport -WorksheetName MigrationReport
            $repor = foreach ($user in $id.identity)
            { 

                if ($id.count -gt 1)
                {
                    $paramWriteProgress = @{
                        Activity        = 'Processing the Mailbox migration Report'
                        Status          = "Processing Samaccount $user  [$i] of [$($id.Count)] users"
                        PercentComplete = (($i / $id.Count) * 100)

                    }

                    Write-Progress @paramWriteProgress
                }
                $i++

    
                Get-EXOMoveRequestStatistics $user | Select-Object DisplayName, Status, batchname, StartTimestamp , SuspendedTimestamp, CompletionTimestamp -ErrorAction silentlycontinue -WarningAction silentlycontinue 
            }




            $repor | export-csv $env:TEMP\report.csv 
 
            import-csv $env:TEMP\report.csv | Export-Excel $filename -AutoSize -TableName MigrationReport -WorksheetName MigrationReport
    
        }
        catch 
        {
            $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
            Write-Host $Errormsg -ForegroundColor magenta
        }
    }
    
    end
    {
    }
}

function connect-complianceCenter
{
    <#
 #... <Short Description>
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.PARAMETER one
    Specifies Pram details.
.PARAMETER two
    Specifies Pram details
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 July 2019         Lawrence       Initial Coding

#>
     
   
    
    
    begin 
    {
        $creds = Get-Credential -Credential lawrence.mckay@urbanutilities.com.au      
    }
    
    process 
    {
            
        try 
        {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $creds -Authentication Basic -AllowRedirection
            Import-PSSession $Session -DisableNameChecking
        }
        catch 
        {
            $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
            Write-Host $Errormsg -ForegroundColor magenta
        }
    }
    
    end
    {
            
    }
}


function get-SyncDetails1
{
    <#
 #... Get if UPN sync'd or not
.SYNOPSIS
    Get if UPN sync'd or not
.DESCRIPTION
    Get if UPN sync'd or not
    This will be user to fix the details to Sync to Azure
.PARAMETER one
    $onpremuser    
    This should be made up by $onpremusers = Get-Mailbox -ResultSize unlimited

.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>get-SyncDetails -onpremusers $users
    C:\PS>get-SyncDetails -onpremusers (Get-Mailbox -ResultSize unlimited)
    
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    20 April 2020
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           20 April 2020         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        $onpremuser   
    )
      
    process 
    {
        $i = 1 
        foreach ($user1 in $onpremuser)
        { 
            $user = get-mailbox $user1
            if ($onpremuser.count -gt 1)
            {
                $paramWriteProgress = @{
                    Activity         = 'Exporting Mailbox Informaiton'
                    Status           = "Processing [$i] of [$($onpremuser.Count)] users"
                    PercentComplete  = (($i / $onpremuser.Count) * 100)
                    CurrentOperation = "Completed : [$user]"
                }
                Write-Progress @paramWriteProgress
            }
            $i++
            # Write-host 'checking if Synced' -ForegroundColor Cyan
            if (get-msoluser -UserPrincipalName $user.UserPrincipalName -erroraction Silentlycontinue) 
            { 
                $synced = 'Exist'
                $OU = $user.OrganizationalUnit
                #       Write-host 'getting AD user info ' -ForegroundColor  Cyan
                $aduser = Get-ADUser -Filter { UserPrincipalName -Eq $user.UserPrincipalName } -Properties * | select givenname, surname, displayname, Enabled, EmailAddress
            
                [PSCustomObject] @{ 
                    UPN                  = $user.UserPrincipalName
                    EmailAddress         = $aduser.EmailAddress
                    Synced               = $synced
                    RecipientTypeDetails = $user.RecipientTypeDetails
                    OU                   = $user.OrganizationalUnit  
                    Firstname            = $aduser.givenname
                    Surname              = $aduser.surname
                    Displayname          = $aduser.displayname
                    AccountEnabled       = $aduser.enabled
                }
            
            }
            else
            {
                $aduser = Get-ADUser -Filter { UserPrincipalName -Eq $user.UserPrincipalName } -Properties * | select givenname, surname, displayname, Enabled, EmailAddress
                $synced = 'Not Synced'
                $OU = $user.OrganizationalUnit   
            }
            [PSCustomObject] @{ 
                UPN                  = $user.UserPrincipalName
                EmailAddress         = $aduser.EmailAddress
                Synced               = $synced
                RecipientTypeDetails = $user.RecipientTypeDetails
                OU                   = $user.OrganizationalUnit  
                Firstname            = $aduser.givenname
                Surname              = $aduser.surname
                Displayname          = $aduser.displayname
                AccountEnabled       = $aduser.enabled
            }
        }
        
    }
       
}
    
Function test-accountSynced
{
  
    <#
 #... Test if user or users synced to AAD
.SYNOPSIS
    Test if user or users synced to AAD
.DESCRIPTION
    Test if user or users synced to AAD
.PARAMETER UPN
    Specifies Pram details.

.EXAMPLE
    C:\PS>test-accountSynced -upn xxx@yyy.net
    C:\PS>test-accountSynced -upn $arrayofUPNs
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    24 April 2020
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           24 April 2020       Lawrence       Initial Coding

#>
      
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'upn')]
        $UPN
    )
    $i = 1
    foreach ($user in $upn)
    {


                
        $paramWriteProgress = @{
            Activity         = 'Getting Sync Status'
            Status           = "Processing [$i] of [$($upn.Count)] users"
            PercentComplete  = (($i / $upn.Count) * 100)
            CurrentOperation = "Completed : [$User]"
        }
        Write-Progress @paramWriteProgress
        $i++

        $azure = 'Not in Azure'
        if (get-msoluser -UserPrincipalName $user -erroraction Silentlycontinue )
        {
            $Azure = 'User is Synced to Azure'
        }
        [PSCustomObject] @{ 
            UPN        = $user
            SyncStatus = $azure
        }

    }
}


Function test-moveExist
{
  
    <#
 #... Test if user or users synced to AAD
.SYNOPSIS
    Test if user or users synced to AAD
.DESCRIPTION
    Test if user or users synced to AAD
.PARAMETER UPN
    Specifies Pram details.

.EXAMPLE
    C:\PS>test-accountSynced -upn xxx@yyy.net
    C:\PS>test-accountSynced -upn $arrayofUPNs
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    24 April 2020
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           24 April 2020       Lawrence       Initial Coding

#>
      
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'upn')]
        $UPN
    )

    $i = 0

    foreach ($user in $upn)
    {
        $move = 'No Move Found'
        if (Get-EXOMoverequest  $user -erroraction Silentlycontinue )
        {
            $move = Get-EXOMoverequest  $user | Select-Object -ExpandProperty Status
            
        }
        [PSCustomObject] @{ 
            UPN        = $user
            MoveStatus = $move
            
        }


        if ($Upn.count -gt 1)
        {
            
            
            $paramWriteProgress = @{
                Activity         = 'checking Mioverequest'
                Status           = "Processing [$i] of [$($Upn.Count)] users"
                PercentComplete  = (($i / $Upn.Count) * 100)
                CurrentOperation = "Completed : [$user]"
            }
            Write-Progress @paramWriteProgress
            
        }
        $i++

    }
}


function remove-allOktaDirectLicenses
{
    #... Removes all Okta groups from Account.
    <#

Syntax
    remove-allOktaDirectLicenses -Samaccountname <account >

Notes

    Neeeds to be run as GA or with an account that has rights to remove licensing as it removes the users licenses directly from AAD.



    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.



    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd

    Date:    15 May 2020


    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           15 May 2020      Lawrence       Initial Coding

   #>
      
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'upns')]
        $Samaccountname
    )

    foreach ($sam in $samaccountname )
    {
        $var = get-aduser $sam
        $upn = $var.UserPrincipalName
        Write-host 'Processing: ' $upn -ForegroundColor Green
        
        try
        { 
            Write-host '    Removing users from all groups that start with OKTAOffice group membership name' -ForegroundColor Cyan
            $365OKTAGroup = Get-ADPrincipalGroupMembership -Identity $sam | where { $_.name -like 'oktaoffice365*' -or $_.name -match 'OKTA-AADSyncOnly_NotLicensed' }

            foreach ($grp in $365OKTAGroup.samaccountname)
            {
                Try
                {
                    Write-host  'Removing:' $sam' from:' $grp -ForegroundColor DarkGreen

                    Remove-ADGroupMember -Identity $grp -Members $var -Confirm:$false 
                }

                catch
                {
                    Write-Output "Failed : $($_.Exception.Message)"
                }

                Write-host '    Removing All OKTA Licensing from user in MSOL' -ForegroundColor Cyan

                (get-MsolUser -UserPrincipalName $upn).licenses.AccountSkuId |
                foreach {
                    Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $_
                }


                Write-host '    Disabling Account' -ForegroundColor Cyan
                Disable-adaccount -Identity $sam
            }
            Catch 
            {
                Write-host "MAjor Error  $($_.Exception.Message)"
            }
        }

        Catch 
        {

        }
    }
}



function get-AllDL_Primary_and_ProxyAddress_with_SAM
{
    <#
 #... Get All DL's Primary and Proxy Addresses
.SYNOPSIS
    Get All Mailbox sufix's
.DESCRIPTION
    gets input of users UPN and then returnes all Email Alias Sufix  Eg after the @ symbol

.PARAMETER UPNs
    UPN or array of users UPN

.PARAMETER showprogress
    Switch to sho progress

    
.EXAMPLE
    C:\PS>get-AllDL_Primary_and_ProxyAddress_with_SAM -UPNs <UPN>
    C:\PS>get-AllDL_Primary_and_ProxyAddress_with_SAM <UPN> -showprogress
    C:\PS>get-AllDL_Primary_and_ProxyAddress_with_SAM -UPNs <UPN> -showprogress |export-csv <file>.csv
    
    

.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           27 May  2019        Lawrence       Initial Coding

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter Users or Array')]
        $name,
        [Switch]$ShowProgress
    
    )

    Begin
    {
        $i = 0
    }
    process 
    {

        foreach ($grp in $name)
        {
            try
            {
                if ($ShowProgress)
                {
                    #sHOW PROGRESS bAR
                    $paramWriteProgress = @{
                        Activity        = "Processing  DL's  "
                        Status          = "Processing  [$i] of [$($name.Count)] DL's"
                        PercentComplete = (($i / $name.Count) * 100)

                    }

                    Write-Progress @paramWriteProgress
                    $i++
                }
                

                $grpsmtps = get-DistributionGroup $grp  -erroraction silentlycontinue  

                foreach ($grpSmtp in $grpsmtps)
                {

                    foreach ($caseSMTP in $grpSmtp.emailaddresses)
                    {
                        $Proxyaddress = $null
                        if ($casesmtp.PrefixString -clike 'smtp')
                        {
                            $Proxyaddress = $casesmtp.smtpaddress
                            [PSCustomObject] @{ 
                                Samaccountname = $grpSmtp.SamAccountName
                                PrimarySMTP    = $grpSmtp.PrimarySmtpAddress
                                Proxyaddress   = $Proxyaddress
                            }
                        }
                    }
                }
            }
            catch
            {
                $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
                Write-Host $Errormsg -ForegroundColor magenta
            }
        }
    }
    end
    {

    }
}


#scripts for database
function get-ADdata
{
    get-aduser -Properties * -Filter * | select AccountExpirationDate	,
    accountExpires	,
    AccountLockoutTime	,
    AccountNotDelegated	,
    adminCount	,
    AllowReversiblePasswordEncryption	,
    AuthenticationPolicy	,
    AuthenticationPolicySilo	,
    BadLogonCount	,
    CannotChangePassword	,
    CanonicalName	,
    Certificates	,
    City	,
    CN	,
    codePage	,
    Company	,
    CompoundIdentitySupported	,
    Country	,
    countryCode	,
    Created	,
    createTimeStamp	,
    Deleted	,
    Department	,
    Description	,
    DisplayName	,
    DistinguishedName	,
    Division	,
    DoesNotRequirePreAuth	,
    dSCorePropagationData	,
    EmailAddress	,
    EmployeeID	,
    EmployeeNumber	,
    Enabled	,
    Fax	,
    GivenName	,
    HomeDirectory	,
    HomedirRequired	,
    HomeDrive	,
    HomePage	,
    HomePhone	,
    Initials	,
    instanceType	,
    isCriticalSystemObject	,
    isDeleted	,
    KerberosEncryptionType	,
    LastBadPasswordAttempt	,
    LastKnownParent	,
    LastLogonDate	,
    LockedOut	,
    LogonWorkstations	,
    Manager	,
    MemberOf	,
    MNSLogonAccount	,
    MobilePhone	,
    Modified	,
    modifyTimeStamp	,
    msDS-User-Account-Control-Computed	,
    Name	,
    nTSecurityDescriptor	,
    ObjectCategory	,
    ObjectClass	,
    ObjectGUID	,
    objectSid	,
    Office	,
    OfficePhone	,
    Organization	,
    OtherName	,
    PasswordExpired	,
    PasswordLastSet	,
    PasswordNeverExpires	,
    PasswordNotRequired	,
    POBox	,
    PostalCode	,
    PrimaryGroup	,
    primaryGroupID	,
    PrincipalsAllowedToDelegateToAccount	,
    ProfilePath	,
    ProtectedFromAccidentalDeletion	,
    pwdLastSet	,
    samAccountName	,
    sAMAccountType	,
    ScriptPath	,
    sDRightsEffective	,
    servicePrincipalName	,
    ServicePrincipalNames	,
    showInAdvancedViewOnly	,
    SID	,
    SIDHistory	,
    SmartcardLogonRequired	,
    State	,
    StreetAddress	,
    Surname	,
    Title	,
    TrustedForDelegation,
    TrustedToAuthForDelegation	,
    UseDESKeyOnly	,
    userAccountControl	,
    userCertificate	,
    UserPrincipalName	,
    uSNChanged	,
    uSNCreated	,
    whenChanged	,
    whenCreated

}

function get-mailboxdata
{
    get-mailbox -ResultSize unlimited | select Database,
    UseDatabaseRetentionDefaults,
    RetainDeletedItemsUntilBackup	,
    DeliverToMailboxAndForward	,
    LitigationHoldEnabled	,
    SingleItemRecoveryEnabled,	
    RetentionHoldEnabled	,
    RetentionPolicy	,
    AddressBookPolicy,	
    CalendarRepairDisabled	,
    ExchangeGuid	,
    ExchangeUserAccountControl	,
    MessageTrackingReadStatusEnabled	,
    ExternalOofOptions	,
    ForwardingAddress	,
    ForwardingSmtpAddress,	
    RetainDeletedItemsFor,	
    IsMailboxEnabled	,
    OfflineAddressBook	,
    ProhibitSendQuota	,
    ProhibitSendReceiveQuota	,
    RecoverableItemsQuota	,
    RecoverableItemsWarningQuota	,
    DowngradeHighPriorityMessagesEnabled	,
    RecipientLimits	,
    IsResource	,
    IsLinked	,
    IsShared	,
    LinkedMasterAccount	,
    ResourceCapacity	,
    ResourceType	,
    SamAccountName	,
    AntispamBypassEnabled	,
    ServerName	,
    UseDatabaseQuotaDefaults	,
    IssueWarningQuota	,
    RulesQuota	,
    Office	,
    UserPrincipalName	,
    UMEnabled	,
    ThrottlingPolicy	,
    RoleAssignmentPolicy	,
    SharingPolicy	,
    ArchiveDatabase	,
    ArchiveGuid	,
    ArchiveQuota,	
    ArchiveWarningQuota	,
    ArchiveDomain	,
    ArchiveStatus	,
    RemoteRecipientType	,
    DisabledArchiveDatabase	,
    QueryBaseDNRestrictionEnabled	,
    MailboxMoveFlags	,
    MailboxMoveStatus	,
    IsPersonToPersonTextMessagingEnabled	,
    IsMachineToPersonTextMessagingEnabled	,
    CalendarVersionStoreDisabled	,
    SKUAssigned	,
    AuditEnabled,	
    AuditLogAgeLimit	,
    WhenMailboxCreated	,
    UsageLocation	,
    HasPicture	,
    HasSpokenName	,
    Alias	,
    ArbitrationMailbox	,
    OrganizationalUnit	,
    CustomAttribute1	,
    CustomAttribute10	,
    CustomAttribute11	,
    CustomAttribute12	,
    CustomAttribute13	,
    CustomAttribute15	,
    CustomAttribute2	,
    DisplayName	,
    ExternalDirectoryObjectId	,
    HiddenFromAddressListsEnabled,	
    MaxSendSize	,
    MaxReceiveSize	,
    ModerationEnabled	,
    EmailAddressPolicyEnabled	,
    PrimarySmtpAddress	,
    RecipientType	,
    RecipientTypeDetails	,
    RequireSenderAuthenticationEnabled	,
    SimpleDisplayName	,
    SendModerationNotifications	,
    WindowsEmailAddress	,
    MailTip	,
    IsValid	,
    ExchangeVersion	,
    Name	,
    Identity	,
    Guid	,
    ObjectCategory	,
    WhenChanged	,
    WhenCreated	,
    WhenChangedUTC	,
    WhenCreatedUTC	,
    OrganizationId	,
    OriginatingServer

}

#End Scripts for Database
Function Publish-MoveStats
{
    #... Generates the MoverequestStats for users.
    <#
    Syntax   Publish-MoveStats -users john.doe@xyz.com
       Syntax   Publish-MoveStats -users #users |export-csv xxxx.csv 

    This will generate the report for all of the users entered


#>


    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $users
    )

    BEGIN
    {
        Write-host 'Generating the Move Request Stats Report'
        $i = 1
    }

    PROCESS
    {

        #Navigate-quuMigrationFolder $batchname
        foreach ($user in $users)
        {
            $moveRequutats = Get-EXOMoveRequest $user | select-object -Expand Identity | Get-EXOMoveRequestStatistics
            

            $moveRequutats | Select-object MailboxIdentity,
            DisplayName,
            ExchangeGUID,
            Status,
            Flags,
            Direction,
            WorkLoadType,
            RecipientTypeDetails,
            SourceServer,
            RemoteHostName,
            BatchName,
            RemoteCredentialUserName,
            TargetDeliveryDomain,
            BadItemLimit,
            BadItemsEncountered,
            LargeItemLimit,
            LargeItemsEncountered,
            QueuedTimestamp,
            StartTimestamp,
            LastUpdateTimestamp,
            LastSuccessfulSyncTimestamp,
            InitialSeedingCompletedTimestamp,
            FinalSyncTimestamp,
            CompletionTimestamp,
            OverallDuration,
            TotalSuspendedDuration,
            TotalFailedDuration,
            TotalQueuedDuration,
            TotalInProgressDuration,
            TotalMailboxSize,
            @{n = 'TotalMailboxSize(MB)'; e = { [math]::Round(($_.TotalMailboxSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) } },
            TotalMailboxItemCount,
            BytesTransferred,
            @{n = 'BytesTransferred(MB)'; e = { [math]::Round(($_.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2) } },
            ItemsTransferred,
            PercentComplete,
            Identity,
            ObjectState 


        }
    }

    END
    {

    }

}


function remove-quuaddressbookpolicy
{
    <#
 #... Sets the the Addressbookpolicy to disabled
.SYNOPSIS
    Sets the the Addressbookpolicy to disabled for the users specified 
.DESCRIPTION
    Sets the the Addressbookpolicy to disabled for the users specified 
.PARAMETER UPNS
    UPN or an array of UPNS
.EXAMPLE
    C:\PS>remove-quuaddressbookpolicy -upn <UPN>
    C:\PS>remove-quuaddressbookpolicy -upn $bulk
    C:\PS>remove-quuaddressbookpolicy -upn $bulk |export-csv c:\temp\addressbook.csv 
        
    Example of how to use this cmdlet
.OUTPUTS
    Output from this cmdlet is in an array form that is outputaable
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2020
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1            15 May 2020         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter UPN here')]
        $UPNS
    )
    
    
    begin 
    {
             
    }
    
    process 
    {
        $i = 1   
      
           
        foreach ($upn in $UPNS)
        {
            #Progress bar.
            $paramWriteProgress = @{
                Activity         = 'Removing the Addressbook policy for Exchange   '
                Status           = "Processing  [$i] of [$($UPNs.Count)] Accounts"
                PercentComplete  = (($i / $UPNs.Count) * 100)
                CurrentOperation = "Completed : [$upn]" 
            }

            Write-Progress @paramWriteProgress
            $i++

            try 
            {
                Set-Mailbox $upn -emailaddresspolicyenabled $false -WarningAction silentlycontinue
                $policyremoved = "Policy Removed Sucessfully"
            }
                    
            catch  
            {
                $Errormsg = 'ERROR : $($.exception.message)'
                Write-Host $Errormsg -ForegroundColor magenta
                $policyremoved = "Policy Removal Failed"
            }

            [PSCustomObject] @{ 
                UPN           = $upn
                policyremoved = $policyremoved

            }

        }
    }
   
    end
    {
                
    }
}

function set-TurnaddresspolicyOFF
{
    <#
 #... Turns off Exchange AddressPolicy for USer
.SYNOPSIS
    Turns off Exchange AddressPolicy for USer
.DESCRIPTION
    Turns off Exchange AddressPolicy for USer / Users   Email address / UPN or Sam
.PARAMETER Users
    can be 1 user or array

.EXAMPLE
    C:\PS>set-QUUaddresspolicyOFF -users xxx@yyy.com   
    C:\PS>set-QUUaddresspolicyOFF -users $bulkusers
Example of how to use this cmdlet

.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Array of output.
    UserDisplayname  SamAcount UserPrincipalName   WindowsEmailAddress    OriginalSettings NewSettings
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 July 2019         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The Output Folder')]
        $users
 
    )
    
    
    begin 
    {
        $output = @()    
    }
    
    process 
    {
      
      
        try 
        {
                

            foreach ($user in $users )
            {
                #Get Original Settings
                $OriginalSettings = get-mailbox $user 

                # turn off the address Policy
                Set-mailbox  $user -EmailAddressPolicyEnabled $false -ErrorAction silentlyContinue -WarningAction SilentlyContinue

                # Get new Settings
                $NewSettings = get-mailbox $user | select-object -ExpandProperty EmailAddressPolicyEnabled -ErrorAction SilentlyContinue -WarningAction SilentlyContinue


                [PSCustomObject] @{ 
                    UserDisplayname     = $OriginalSettings.DisplayName                
                    SamAcount           = $OriginalSettings.samaccountname
                    UserPrincipalName   = $OriginalSettings.UserPrincipalName
                    WindowsEmailAddress = $OriginalSettings.WindowsEmailAddress
                    OriginalSettings    = $OriginalSettings.EmailAddressPolicyEnabled
                    NewSettings         = $NewSettings
                }
                    
            }
        }
        catch 
        {
            $Errormsg = 'ERROR : $($MyInvocation.InvocationName) `t`t $($_.exception.message)'
            Write-Host $Errormsg -ForegroundColor magenta
        }
        
    
        
    }
    end
    {
            
    }
}


Function Start-cleanUPterminatedusers
{

     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $bulk
    )
    $i = 1
    
    foreach ($Samaccount in $bulk)
    {

        
        if ($Bulk.count -gt 1)
        {
            $paramWriteProgress = @{
                Activity         = 'Cleaning up Terminated users.'
                Status           = "Processing [$i] of [$($Bulk.Count)] users"
                PercentComplete  = (($i / $Bulk.Count) * 100)
                CurrentOperation = "Completed : [$Samaccount]"
            }
            Write-Progress @paramWriteProgress
        }
                
        $i++


        #rRemove all distrubution group membership.
        $ou = get-aduser  $Samaccount -Properties displayname, givenname, Surname, DistinguishedName, canonicalName
        try
        {
            $global:dlsremove = "No DL's Removed'"
            # Get-ADUser $user |
            $DLs = Get-ADPrincipalGroupMembership -Identity $($Samaccount) | Select-Object -Expand Distinguishedname | Get-DistributionGroup -EA 0
            foreach ($DL in $DLS.name)
            {
                ## commented out so not to Run    
                Remove-DistributionGroupMember -Identity $dl -Member  $Samaccount -Confirm:$false -ErrorAction silentlycontinue
            }
            $global:dlsremove = $dls
            [string]$DLREMOVE =$dls

        }
        catch
        {
            $global:dlsremove = 'Error removing User from DLs'
                  
        }
             
             
        ## Hide user from Addressbook.
                
        try
        {
            $global:HideAddressbook = 'User is Already hidden from the Addressbook'
            ## Gets user Status from AD to see if hidden from Addressbook
            $hidden = get-aduser $($Samaccount) -Properties msExchHideFromAddressLists | Select-Object MsExchHideFromAddressLists
            if ($hidden -notmatch "True")
            {
                ##  Hides user from Addressbook
                set-aduser $($Samaccount) -add @{msExchHideFromAddressLists = "TRUE" }                  
                ##Set output for Final results.
                $global:HideAddressbook = 'User is hidden from the Addressbook'
            }
        }
        catch
        {
            $global:HideAddressbook = 'User is Already hidden from the Addressbook'
                  
        }
                

        ##  Remove Office365 Groups
                
        try
        {
            $global:365OKTAGroup = 'No Office365 groups to remove..'
            $global:365OKTAGroup = Get-ADPrincipalGroupMembership -Identity $($Samaccount) | Where-Object { $_.name -like 'oktaoffice365*' -or $_.name -match 'OKTA-AADSyncOnly_NotLicensed' }
            [string]$offgroup = 365OKTAGroup.name
            foreach ($grp in $365OKTAGroup.samaccountname)
            {
                Try
                {
                    ## commented out so not to Run    
                    Remove-ADGroupMember -Identity $grp -Members $Samaccount -Confirm:$false 
                }

                catch
                {
                    #Write-Output 'Failed : $error[0]'
                    $global:365OKTAGroup = 'No Office365 groups to remove..'
                }
                

            }
            
        }
        catch 
        {
  
                    
        }


        ##  Remove Office365 Groups
        #move the use to the OU below.
        try 
        {
            
            $ou2move = "OU=Terminated Users, OU=Users, OU=QUU, DC=corporate, DC=urbanutilities, DC=internal"
            
            Move-ADObject $ou.DistinguishedName -TargetPath $OU2move
            $oumove = =$ou 
        
        }
        catch 
        {
            $oumove = $error[0]
        }
    
        #remove-msoluser -UserPrincipalName $ou.UserPrincipalName -
                
        [PSCustomObject] @{ 
            OU                         = $ou.canonicalName 
            Displayname                = $ou.Displayname
            Firstname                  = $ou.Givenname
            Surname                    = $ou.Surname
            'Sam / Payroll Number'     = $ou.Samaccountname
            UPN                        = $ou.userprincipalname
            'Office365 groups removed' = $offgroup
            'DLs removed'              = $DLREMOV
            'Hidden from Addressbook'  = $global:HideAddressbook
            'accountMoveTo'            = $ou2move
        }
    }
}

Function Start-comparedls
{
    <#
 #... Compares Membership onprem DL's with EXO and AAD groups
.SYNOPSIS
    Compares Membership onprem DL's with EXO and AAD groups
.DESCRIPTION
    Compares Membership onprem DL's with EXO and AAD groups
.PARAMETER DLs
    Imput of DL's NAmes 

.EXAMPLE
    C:\PS>start-dlcompare 
    C:\PS>start-dlcompare -DLs <dlname>

    Example of how to use this cmdlet

    .INPUTS
    Inputs to this cmdlet (if any)

    .OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 May 2019
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           14 Oct 2020         Lawrence       Initial Coding

#>
    [CmdletBinding()]
    param( [Parameter(Mandatory = $false,
            HelpMessage = 'Enter The DL to check')]
        $DLS = (Get-EXODistributionGroup -resultsize unlimited | select -ExpandProperty Identity)
    )

    #$DLS = Get-EXODistributionGroup -resultsize unlimited | select -ExpandProperty Identity
    begin
    {
    }
    process
    {
        $i = 1
        foreach ($name in $DLS)
        {
            #sHOW PROGRESS bAR
            $paramWriteProgress = @{
                Activity        = "Processing  DL $name   "
                Status          = "Processing  [$i] of [$($DLS.Count)] DLS"
                PercentComplete = (($i / $DLs.Count) * 100)

            }

            Write-Progress @paramWriteProgress
            $i++

            $id = Get-exoDistributionGroup $($name) | select -ExpandProperty ExternalDirectoryObjectId
            $aad = (Get-MsolGroupMember -GroupObjectId $id -erroraction silentlycontinue ).count
            $EXODL = (Get-exoDistributionGroupMember $($name) -erroraction silentlycontinue ).count
            $ONPREM = (Get-DistributionGroupMember $($name) -erroraction silentlycontinue ).count
            $matched = 'match'
            $Onpremmatch = 'match'

            if ($aad -ne $exodl) { $matched = "notmatchedAAD2EXO" }
            if ($ONPREM -ne $aad) { $Onpremmatch = "notmatchedOnPrem2AAD" }


            [PSCustomObject] @{ 
                Groupname          = $name
                AADGroupmembership = $aad
                EXODLMembership    = $EXODL
                OnpremCount        = $ONPREM
                MatchedCloud       = $matched
                MatchedOnprem      = $Onpremmatch
            }

        }
    }

}



function check-movesyncstatus
{
    [CmdletBinding()]
    param
    (
        $Bulk
       
       
    )

    $i = 1
    foreach ($mb in $bulk)
    {
        $Details = ''

        #sHOW PROGRESS bAR
        $paramWriteProgress = @{
            Activity        = 'Processing  Mmoves $mb  '
            Status          = "Processing  [$i] of [$($bulk.Count)] Mailboxes"
            PercentComplete = (($i / $bulk.Count) * 100)

        }

        Write-Progress @paramWriteProgress
        $i++

        try
        {    
            $dn = get-aduser -filter { userPrincipalName -eq $mb } -ErrorAction silentlycontinue | select -ExpandProperty DistinguishedName
            $move = get-exomoverequest $mb -erroraction silentlycontinue
            $status = $move.Status
            $DName = $move.displayname
            
            
            if ($null -eq $move )
            {
                $Details = 'no move found'
        
            }
        
        
        }
        catch
        {
            $move = ''
            $Details = 'no movefound'


        }
 
        [PSCustomObject] @{ 
            Name              = $mb
            Displayname       = $dname
            Status            = $status
            details           = $Details
            DistinguishedName = $dn

        }
    }
}

Function Compare-targetaddress
{
    
    #... Checks to see if targetaddress is in proxyaddress 
    <#
.SYNOPSIS
    Checks to see if targetaddress is in proxyaddress 
.DESCRIPTION
    This function will compare the Target address against the Proxy Address in Local Ad and AzureAD.  ThThe target address  MUST be in the Proxyaddress otherwise autodiscovery does not work.
    Note  
        AD Powershell module is required to be installed
   
	
.PARAMETER BulkUsers
    This can be a single users SamAccountName or an array 
	
.EXAMPLE
        Compare-targetaddress -bulkUsers $users
        Compare-targetaddress -bulkUsers 123456
        Compare-targetaddress -bulkUsers $users |out-gridview
        Compare-targetaddress -bulkUsers $users |export-csv -NoTypeInformation -Path c:\tmp\immutableID.csv    Example of how to use this cmdlet

.INPUTS
    BulkUsers
	
.OUTPUTS
    This will output the following 
        Matched 
        DisplayName  
        SamAccountName            
        UPN                      
        OnpremTargetaddress  
        DN

.NOTES
        Needs connection to AD 

          

    *******************
    Copyright Notice.
    *******************
    Copyright (c) 2018   McKayIT Solutions Pty Ltd.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.



    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd

    Date:    14 Sept 2018


    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           12 Nov 2018         Lawrence       Initial Coding

 #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String[]]$bulkUsers

    )

    begin
    {
        $i = 1
    }

    Process
    {

        foreach ($line in $bulkUsers)
        {

            ## progress bar

            if ($Bulk.count -gt 1)
            {
                $paramWriteProgress = @{
                    Activity        = 'Cleaning up Terminated users.'
                    Status          = "Processing [$i] of [$($Bulk.Count)] users"
                    PercentComplete = (($i / $Bulk.Count) * 100)
                    
                }
                Write-Progress @paramWriteProgress
            }
                
            $i++





            $onprem = get-aduser $line  -Properties *

            
            if ($onprem.proxyaddresses -contains $onprem.targetaddress)
            {
                #        write-host $online.userprincipalName ' is the same both on Prem and Online' -ForegroundColor green
                $compared = 'Same'

            }
            else
            {
                $compared = 'Different'
            }

            [PSCustomObject] @{
                Matched             = $compared
                DisplayName         = $onprem.Displayname
                SamAccountName      = $onprem.samaccountname
                UPN                 = $onprem.UserPrincipalName
                OnpremTargetaddress = $onprem.targetaddress
                DN                  = $onprem.DistinguishedName
            }

        }

    }
    END
    {

    }
}


function export-MoveinfoPostAndPre
{
    <#
 #... Dump Move Report
.SYNOPSIS
    Dump Move Report post and Pre Migration
.DESCRIPTION
  Dump Move Report post and Pre Migration

  MoveDate is optional as it will put the current date in if not Entered as a Var.
    
.PARAMETER Emailaddress
    USer / Users in an array

.PARAMETER filename
    path and filename to save Excel file to 
    eg c:\tmp\first.xlsx

.PARAMETER MigrationDate
    This will auto add the Current date but You can also manually enter the Date in the Following format 'YYYY-MM-DD"



.EXAMPLE
    C:\PS>gexport-Moveinfo -Emailaddress <Emailaddress> 
    C:\PS>gexport-Moveinfo -Emailaddress <Emailaddress>  -migrationdate '2020-10-28'
       
    


.NOTES

Needs a connection to Exchange, AD and MSOL Service

    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    10 August  2020
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           10 Aug 2020         Lawrence       Initial Coding

#>


    param
    (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter UPN here')]
        $Emailaddress,
        $MigrationDate = (get-date  -format yyyy-MM-dd)
    )

    Try
    {
        $i = 1   
        foreach ($user in $Emailaddress)
        {

            #Progress bar.
            $paramWriteProgress = @{
                Activity         = 'Creating Move Report   '
                Status           = "Processing  [$i] of [$($Emailaddress.Count)] Accounts"
                PercentComplete  = (($i / $Emailaddress.Count) * 100)
                CurrentOperation = "Processed: [$user]" 
            }

            Write-Progress @paramWriteProgress
            $i++



            $Userinfo = Get-Recipient $user 
            $Migrated = [bool](Get-exomailbox $user -erroraction SilentlyContinue)
            $CorpMobileUser = 'False'
            $grpmem = get-adgroupmembership $user
            If ($grpmem -match 'Mobile')
            {
                $CorpMobileUser = 'True'
            }
      
            [PSCustomObject] @{ 
                Displayname        = $Userinfo.Displayname
                SamAccountName     = $Userinfo.Alias
                PrimarySMTPAddress = $userinfo.PrimarySmtpAddress
                MAilboxType        = $Userinfo.RecipientTypeDetails
                MigrationDate      = $migrationDate
                HasCorpMobile      = $CorpMobileUser
                Migrated           = $migrated
            }
  
        }
    }
    catch
    {
        Write-Host $Error[0] -ForegroundColor Magenta
    }
}

$Bulk = gc  C:\tmp\22nd.txt
$filename = 'c:\tmp\2020-10-22-pre.xlsx'
#  date format  YYYY=MM-DD
$migrationDate = '2020-10-22'
 export-MoveinfoPostAndPre -Emailaddress $Bulk  | export-excel  $filename -AutoSize -TableName Migration -WorksheetName Migration

