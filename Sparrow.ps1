Function Check-PSModules{

    [cmdletbinding()]Param()

    $ModuleArray = @("CloudConnect","AzureAD","MSOnline")

    ForEach ($ReqModule in $ModuleArray){
        If ($null -eq (Get-Module $ReqModule -ListAvailable -ErrorAction SilentlyContinue)){
            Write-Verbose "Required module, $ReqModule, is not installed on the system."
            Write-Verbose "Installing $ReqModule from default repository"
            Install-Module -Name $ReqModule -Force
            Write-Verbose "Importing $ReqModule"
            Import-Module -Name $ReqModule
        } ElseIf ($null -eq (Get-Module $ReqModule -ErrorAction SilentlyContinue)){
            Write-Verbose "Importing $ReqModule"
            Import-Module -Name $ReqModule
        }
    }

    #If you want to change the default export directory, please change the $ExportDir value.
    #Otherwise, the default export is the user's home directory, Desktop folder, and ExportDir folder.
    $ExportDir = "$home\Desktop\ExportDir"
    If (!(Test-Path $ExportDir)){
        New-Item -Path $ExportDir -ItemType "Directory" -Force
    }
}

Function Get-UALData {

    [cmdletbinding()]Param()

    #Calling on CloudConnect to connect to the tenant's Exchange Online environment via PowerShell
    Connect-EXO

    $EndDate = (Get-Date)
    #If you want to change the length of time, please edit the value in $StartDate accordingly
    #Example, if you want to only pull 45 days, the $StartDate value would be (Get-Date).AddDays(-45)
    $StartDate = (Get-Date).AddDays(-90)

    $LicenseQuestion = Read-Host 'Do you have an Office 365/Microsoft 365 E5/G5 license? Y/N'
    Switch ($LicenseQuestion){
        Y {$LicenseAnswer = "Yes"}
        N {$LicenseAnswer = "No"}
    }
    $AppIdQuestion = Read-Host 'Would you like to investigate a certain application? Y/N'
    Switch ($AppIdQuestion){
        Y {$AppIdInvestigation = "Yes"}
        N {$AppIdInvestigation = "No"}
    }
    If ($AppIdInvestigation -eq "Yes"){
        $SusAppId = Read-Host "Enter the application's AppID to investigate"
    } Else{
        Write-Host "Skipping AppID investigation"
    }
    
    #Searches for any modifications to the domain and federation settings on a tenant's domain
    Write-Verbose "Searching for 'Set domain authentication' and 'Set federation settings on domain' operations in the UAL."
    $DomainData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Set domain authentication","Set federation settings on domain" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as Domain_Operations_Export.csv
    Export-UALData -UALInput $DomainData -CsvName "Domain_Operations_Export" -WorkloadType "AAD"

    #Searches for any modifications or credential modifications to an application
    Write-Verbose "Searching for 'Update application' and 'Update application ? Certificates and secrets management' in the UAL."
    $AppData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Update application","Update application ? Certificates and secrets management" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as AppUpdate_Operations_Export.csv
    Export-UALData -UALInput $AppData -CsvName "AppUpdate_Operations_Export" -WorkloadType "AAD"

    #Searches for any modifications or credential modifications to a service principal
    Write-Verbose "Searching for 'Update service principal' and 'Add service principal credentials' in the UAL."
    $SpData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Update service principal","Add service principal credentials" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as ServicePrincipal_Operations_Export.csv   
    Export-UALData -UALInput $SpData -CsvName "ServicePrincipal_Operations_Export" -WorkloadType "AAD"

    #Searches for any app role assignments to service principals, users, and groups
    Write-Verbose "Searching for 'Add app role assignment to service principal', 'Add app role assignment grant to user', and 'Add app role assignment to group' in the UAL."
    $AppRoleData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Add app role assignment" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as AppRoleAssignment_Operations_Export.csv      
    Export-UALData -UALInput $AppRoleData -CsvName "AppRoleAssignment_Operations_Export" -WorkloadType "AAD"

    #Searches for any OAuth or application consents
    Write-Verbose "Searching for 'Add OAuth2PermissionGrant' and 'Consent to application' in the UAL."
    $ConsentData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Add OAuth2PermissionGrant","Consent to application" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as Consent_Operations_Export.csv       
    Export-UALData -UALInput $ConsentData -CsvName "Consent_Operations_Export" -WorkloadType "AAD"

    #Searches for SAML token usage anomaly (UserAuthenticationValue of 16457) in the Unified Audit Logs
    Write-Verbose "Searching for 16457 in UserLoggedIn and UserLoginFailed operations in the UAL."
    $SAMLData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "UserLoggedIn","UserLoginFailed" -ResultSize 5000 -FreeText "16457" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as SAMLToken_Operations_Export.csv      
    Export-UALData -UALInput $SAMLData -CsvName "SAMLToken_Operations_Export" -WorkloadType "AAD"

    #Searches for PowerShell logins into mailboxes
    Write-Verbose "Searching for PowerShell logins into mailboxes in the UAL."
    $PSMailboxData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 -Operations "MailboxLogin" -FreeText "Powershell" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as PSMailbox_Operations_Export.csv      
    Export-UALData -UALInput $PSMailboxData -CsvName "PSMailbox_Operations_Export" -WorkloadType "EXO2"

    #Searches for well-known AppID for Exchange Online PowerShell
    Write-Verbose "Searching for PowerShell logins using known PS application ids in the UAL."
    $PSLoginData1 = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000  -FreeText "a0c73c16-a7e3-4564-9a95-2bdf47383716" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as PSLogin_Operations_Export.csv  
    Export-UALData -UALInput $PSLoginData1 -CsvName "PSLogin_Operations_Export" -WorkloadType "AAD"

    #Searches for well-known AppID for PowerShell
    $PSLoginData2 = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000  -FreeText "1b730954-1685-4b74-9bfd-dac224a7b894" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #The resultant CSV will be appended with the $PSLoginData* resultant CSV.
    #If you want a separate CSV with a different name, remove the -AppendType parameter (-AppendType "Append")
    #By default, it will show up as part of the PSLogin_Operations_Export.csv  
    Export-UALData -UALInput $PSLoginData2 -CsvName "PSLogin_Operations_Export" -WorkloadType "AAD" -AppendType "Append"

    #Searches for WinRM useragent string in the user logged in and user login failed operations
    $PSLoginData3 = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 -Operations "UserLoggedIn","UserLoginFailed" -FreeText "WinRM" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #The resultant CSV will be appended with the $PSLoginData* resultant CSV.
    #If you want a separate CSV with a different name, remove the -AppendType parameter (-AppendType "Append")
    #By default, it will show up as part of the PSLogin_Operations_Export.csv 
    Export-UALData -UALInput $PSLoginData3 -CsvName "PSLogin_Operations_Export" -WorkloadType "AAD" -AppendType "Append"

    If ($AppIdInvestigation -eq "Yes"){
        If ($LicenseAnswer -eq "Yes"){
            #Searches for the AppID to see if it accessed mail items.
            Write-Verbose "Searching for $SusAppId in the MailItemsAccessed operation in the UAL."
            $SusMailItems = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "MailItemsAccessed" -ResultSize 5000 -FreeText $SusAppId -Verbose | Select-Object -ExpandProperty AuditData | Convertfrom-Json
            #You can modify the resultant CSV output by changing the -CsvName parameter
            #By default, it will show up as MailItems_Operations_Export.csv  
            Export-UALData -UALInput $SusMailItems -CsvName "MailItems_Operations_Export" -WorkloadType "EXO"
        } else {
             Write-Host "MailItemsAccessed query will be skipped as it is not present without an E5/G5 license."
        }

        #Searches for the AppID to see if it accessed Sharepoint or OneDrive items
        Write-Verbose "Searching for $SusAppId in the FileAccessed and FileAccessedExtended operations in the UAL."
        $SusFileItems = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "FileAccessed","FileAccessedExtended" -ResultSize 5000 -FreeText $SusAppId -Verbose | Select-Object -ExpandProperty AuditData | Convertfrom-Json
        #You can modify the resultant CSV output by changing the -CsvName parameter
        #By default, it will show up as FileItems_Operations_Export.csv  
        Export-UALData -UALInput $SusFileItems -CsvName "FileItems_Operations_Export" -WorkloadType "Sharepoint"
    }
}

Function Get-AzureDomains{

    [cmdletbinding()]Param()

    #Connect to AzureAD
    Connect-AzureAD

    $DomainData = Get-AzureADDomain
    $DomainArr = @()
    
    ForEach ($Domain in $DomainData){
        $DomainProps = [ordered]@{
            AuthenticationType = $Domain.AuthenticationType
            AvailabilityStatus = $Domain.AvailabilityStatus
            ForceDeleteState = $Domain.ForceDeleteState
            IsAdminManaged = $Domain.IsAdminManaged
            IsDefault = $Domain.IsDefault
            IsInitial = $Domain.IsInitial
            IsRoot = $Domain.IsRoot
            IsVerified = $Domain.IsVerified
            Name = $Domain.Name
            State = $Domain.State
            SupportedServices = ($Domain.SupportedServices -join ';')
        }
        $DomainObj = New-Object -TypeName PSObject -Property $DomainProps
        $DomainArr += $DomainObj
    }
    $DomainArr | Export-Csv $home\Desktop\ExportDir\Domain_List.csv -NoTypeInformation
}

Function Get-AzureSPAppRoles{

    [cmdletbinding()]Param()

    #Connect to your tenant's AzureAD environment
    Connect-AzureAD

    #Retrieve all service principals that are applications
    $SPArr = Get-AzureADServicePrincipal -All $true | Where-Object {$_.ServicePrincipalType -eq "Application"}

    #Retrieve all service principals that have a display name of Microsoft Graph
    $GraphSP = Get-AzureADServicePrincipal -All $true | Where-Object {$_.DisplayName -eq "Microsoft Graph"}

    $GraphAppRoles = $GraphSP.AppRoles | Select -Property AllowedMemberTypes, Id, Value

    $AppRolesArr = @()
    Foreach ($SP in $SPArr) {
        $GraphResource = Get-AzureADServiceAppRoleAssignedTo -ObjectId $SP.ObjectId | Where-Object {$_.ResourceDisplayName -eq "Microsoft Graph"}
        ForEach ($GraphObj in $GraphResource){
            For ($i=0; $i -lt $GraphAppRoles.Count; $i++){
                if ($GraphAppRoles[$i].Id -eq $GraphObj.Id) {
                    $ListProps = [ordered]@{
                        ApplicationDisplayName = $GraphObj.PrincipalDisplayName
                        ClientID = $GraphObj.PrincipalId
                        Value = $GraphAppRoles[$i].Value
                    }
                }
            }
            $ListObj = New-Object -TypeName PSObject -Property $ListProps
            $AppRolesArr += $ListObj 
            }
        }
    #If you want to change the default export directory, please change the $home\Desktop\ExportDir value.
    #Otherwise, the default export is the user's home directory, Desktop folder, and ExportDir folder.
    #You can change the name of the CSV as well, the default name is "ApplicationGraphPermissions"
    $AppRolesArr | Export-Csv $home\Desktop\ExportDir\ApplicationGraphPermissions.csv -NoTypeInformation
}

Function Export-UALData {
    Param(
        [Parameter(ValueFromPipeline=$True)]
        [Object[]]$UALInput,
        [Parameter()]
        [String]$CsvName,
        [Parameter()]
        [String]$WorkloadType,
        [Parameter()]
        [String]$AppendType
        )

        $DataArr = @()
        If ($WorkloadType -eq "AAD") {
            ForEach ($Data in $UALInput){
                $DataProps = [ordered]@{
                    CreationTime = $Data.CreationTime
                    Id = $Data.Id
                    Operation = $Data.Operation
                    Organization = $Data.Organization
                    RecordType = $Data.RecordType
                    ResultStatus = $Data.ResultStatus
                    LogonError = $Data.LogonError
                    UserKey = $Data.UserKey
                    UserType = $Data.UserType
                    Version = $Data.Version
                    Workload = $Data.Workload
                    ClientIP = $Data.ClientIP
                    ObjectId = $Data.ObjectId
                    UserId = $Data.UserId
                    AzureActiveDirectoryEventType = $Data.AzureActiveDirectoryEventType
                    ExtendedProperties = ($Data.ExtendedProperties | ConvertTo-Json -Compress | Out-String).Trim()
                    ModifiedProperties = (($Data.ModifiedProperties | ConvertTo-Json -Compress) -replace "\\r\\n" | Out-String).Trim()
                    Actor = ($Data.Actor | ConvertTo-Json -Compress | Out-String).Trim()
                    ActorContextId = $Data.ActorContextId
                    ActorIpAddress = $Data.ActorIpAddress
                    InterSystemsId = $Data.InterSystemsId
                    IntraSystemId = $Data.IntraSystemId
                    SupportTicketId = $Data.SupportTicketId
                    Target = ($Data.Target | ConvertTo-Json -Compress | Out-String).Trim()
                    TargetContextId = $Data.TargetContextId
                }
                $DataObj = New-Object -TypeName PSObject -Property $DataProps
                $DataArr += $DataObj           
            }
        } elseif ($WorkloadType -eq "EXO"){
            ForEach ($Data in $UALInput){
                $DataProps = [ordered]@{
                    CreationTime = $Data.CreationTime
                    Id = $Data.Id
                    Operation = $Data.Operation
                    OrganizationId = $Data.OrganizationId
                    RecordType = $Data.RecordType
                    ResultStatus = $Data.ResultStatus
                    UserKey = $Data.UserKey
                    UserType = $Data.UserType
                    Version = $Data.Version
                    Workload = $Data.Workload
                    UserId = $Data.UserId
                    AppId = $Data.AppId
                    ClientAppId = $Data.ClientAppId
                    ClientIPAddress = $Data.ClientIPAddress
                    ClientInfoString = $Data.ClientInfoString
                    ExternalAccess = $Data.ExternalAccess
                    InternalLogonType = $Data.InternalLogonType
                    LogonType = $Data.LogonType
                    LogonUserSid = $Data.LogonUserSid
                    MailboxGuid = $Data.MailboxGuid
                    MailboxOwnerSid = $Data.MailboxOwnerSid
                    MailboxOwnerUPN = $Data.MailboxOwnerUPN
                    OperationProperties = ($Data.OperationProperties | ConvertTo-Json -Compress | Out-String).Trim()
                    OrganizationName = $Data.OrganizationName
                    OriginatingServer = $Data.OriginatingServer
                    Folders = ((($Data.Folders | ConvertTo-Json -Compress).replace("\u003c","")).replace("\u003e","")  | Out-String).Trim()
                    OperationCount = $Data.OperationCount
                }
                $DataObj = New-Object -TypeName PSObject -Property $DataProps
                $DataArr += $DataObj           
            }
        } elseif ($WorkloadType -eq "EXO2"){
            ForEach ($Data in $UALInput){
                $DataProps = [ordered]@{
                    CreationTime = $Data.CreationTime
                    Id = $Data.Id
                    Operation = $Data.Operation
                    OrganizationId = $Data.OrganizationId
                    RecordType = $Data.RecordType
                    ResultStatus = $Data.ResultStatus
                    UserKey = $Data.UserKey
                    UserType = $Data.UserType
                    Version = $Data.Version
                    Workload = $Data.Workload
                    ClientIP = $Data.ClientIP
                    UserId = $Data.UserId
                    ClientIPAddress = $Data.ClientIPAddress
                    ClientInfoString = $Data.ClientInfoString
                    ExternalAccess = $Data.ExternalAccess
                    InternalLogonType = $Data.InternalLogonType
                    LogonType = $Data.LogonType
                    LogonUserSid = $Data.LogonUserSid
                    MailboxGuid = $Data.MailboxGuid
                    MailboxOwnerSid = $Data.MailboxOwnerSid
                    MailboxOwnerUPN = $Data.MailboxOwnerUPN
                    OrganizationName = $Data.OrganizationName
                    OriginatingServer = $Data.OriginatingServer
                }
                $DataObj = New-Object -TypeName PSObject -Property $DataProps
                $DataArr += $DataObj           
            }
        } elseif ($WorkloadType -eq "Sharepoint"){
            ForEach ($Data in $UALInput){
                $DataProps = [ordered]@{
                    CreationTime = $Data.CreationTime
                    Id = $Data.Id
                    Operation = $Data.Operation
                    OrganizationId = $Data.OrganizationId
                    RecordType = $Data.RecordType
                    UserKey = $Data.UserKey
                    UserType = $Data.UserType
                    Version = $Data.Version
                    Workload = $Data.Workload
                    ClientIP = $Data.ClientIP
                    ObjectId = $Data.ObjectId
                    UserId = $Data.UserId
                    ApplicationId = $Data.ApplicationId
                    CorrelationId = $Data.CorrelationId
                    EventSource = $Data.EventSource
                    ItemType = $Data.ItemType
                    ListId = $Data.ListId
                    ListItemUniqueId = $Data.ListItemUniqueId
                    Site = $Data.Site
                    UserAgent = $Data.UserAgent
                    WebId = $Data.WebId
                    HighPriorityMediaProcessing = $Data.HighPriorityMediaProcessing
                    SourceFileExtension = $Data.SourceFileExtension
                    SiteUrl = $Data.SiteUrl
                    SourceFileName = $Data.SourceFileName
                    SourceRelativeUrl = $Data.SourceRelativeUrl
                }
                $DataObj = New-Object -TypeName PSObject -Property $DataProps
                $DataArr += $DataObj
            }
        }
        If ($AppendType -eq "Append"){
            $DataArr | Export-csv $home\Desktop\ExportDir\$CsvName.csv -NoTypeInformation -Append
        } Else {
            $DataArr | Export-csv $home\Desktop\ExportDir\$CsvName.csv -NoTypeInformation
        }
        
        Remove-Variable UALInput -ErrorAction SilentlyContinue
        Remove-Variable Data -ErrorAction SilentlyContinue
        Remove-Variable DataObj -ErrorAction SilentlyContinue
        Remove-Variable DataProps -ErrorAction SilentlyContinue
        Remove-Variable DataArr -ErrorAction SilentlyContinue
}

#Function calls, if you do not need a particular check, you can comment it out below with #
Check-PSModules -Verbose
Get-UALData -Verbose
Get-AzureDomains -Verbose
Get-AzureSPAppRoles -Verbose
