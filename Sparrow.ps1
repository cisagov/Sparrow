[cmdletbinding()]Param(
    [Parameter()]
    [string] $AzureEnvironment,
    [Parameter()]
    [string] $ExchangeEnvironment,
    [Parameter()]
    [datetime] $StartDate = [DateTime]::UtcNow.AddDays(-364),
    [Parameter()]
    [datetime] $EndDate = [DateTime]::UtcNow,
    [Parameter()]
    [string] $ExportDir = (Join-Path ([Environment]::GetFolderPath("Desktop")) 'ExportDir'),
    [Parameter()]
    [switch] $NoO365 = $false
)

Function Import-PSModules{

    [cmdletbinding()]Param(
        [Parameter(Mandatory=$true)]
        [string] $ExportDir
        )

    $ModuleArray = @("ExchangeOnlineManagement","AzureAD","MSOnline")

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
    If (!(Test-Path $ExportDir)){
        New-Item -Path $ExportDir -ItemType "Directory" -Force
    }
}

Function Get-AzureEnvironments() {

    [cmdletbinding()]Param(
        [Parameter()]
        [string] $AzureEnvironment, 
        [Parameter()]
        [string] $ExchangeEnvironment
        )

    $AzureEnvironments = [Microsoft.Open.Azure.AD.CommonLibrary.AzureEnvironment]::PublicEnvironments.Keys
    While ($AzureEnvironments -cnotcontains $AzureEnvironment -or [string]::IsNullOrWhiteSpace($AzureEnvironment)) {
        Write-Host 'Azure Environments'
        Write-Host '------------------'
        $AzureEnvironments | ForEach-Object { Write-Host $_ }
        $AzureEnvironment = Read-Host 'Choose your Azure Environment [AzureCloud]'
        If ([string]::IsNullOrWhiteSpace($AzureEnvironment)) { $AzureEnvironment = 'AzureCloud' }
    }

    If ($NoO365 -eq $false) {
        $ExchangeEnvironments = [System.Enum]::GetNames([Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment])
        While ($ExchangeEnvironments -cnotcontains $ExchangeEnvironment -or [string]::IsNullOrWhiteSpace($ExchangeEnvironment) -and $ExchangeEnvironment -ne "None") {
            Write-Host 'Exchange Environments'
            Write-Host '---------------------'
            $ExchangeEnvironments | ForEach-Object { Write-Host $_ }
            Write-Host 'None'
            $ExchangeEnvironment = Read-Host 'Choose your Exchange Environment [O365Default]'
            If ([string]::IsNullOrWhiteSpace($ExchangeEnvironment)) { $ExchangeEnvironment = 'O365Default' }
        }
    } Else {
        $ExchangeEnvironment = "None"
    }

    Return ($AzureEnvironment, $ExchangeEnvironment)
}

Function New-ExcelFromCsv() {

    [cmdletbinding()]Param(
        [Parameter(Mandatory=$true)]
        [string] $ExportDir
        )

    Try {
        $Excel = New-Object -ComObject Excel.Application
    }
    Catch { 
        Write-Host 'Warning; Excel not found - skipping combined file.' 
        Return
    }

    #Open each file and move it in a single workbook
    $Excel.DisplayAlerts = $False
    $Workbook = $Excel.Workbooks.Add()
    $Csvs = Get-ChildItem -Path "${ExportDir}\*.csv" -Force
    $ToDeletes = $Workbook.Sheets | Select-Object -ExpandProperty Name
    ForEach ($Csv in $Csvs) {
        $TempWorkbook = $Excel.Workbooks.Open($Csv.FullName)
        $TempWorkbook.Sheets[1].Copy($Workbook.Sheets[1], [Type]::Missing) | Out-Null
        $Workbook.Sheets[1].UsedRange.Columns.AutoFit() | Out-Null
        $Workbook.Sheets[1].Name = $Csv.BaseName -replace '_Operations_.*',''
    }

    #Save out the new file
    ForEach ($ToDelete in $ToDeletes) { 
        $Workbook.Activate()
        $Workbook.Sheets[$ToDelete].Activate()
        $Workbook.Sheets[$ToDelete].Delete()
    }
    $Workbook.SaveAs((Join-Path $ExportDir 'Summary_Export.xlsx'))
    $Excel.Quit()
}

Function Get-UALData {

    [cmdletbinding()]Param(
        [Parameter(Mandatory=$true)]
        [datetime] $StartDate,
        [Parameter(Mandatory=$true)]
        [datetime] $EndDate,
        [Parameter(Mandatory=$true)]
        [string] $AzureEnvironment,
        [Parameter(Mandatory=$true)]
        [string] $ExchangeEnvironment,
        [Parameter(Mandatory=$true)]
        [string] $ExportDir
        )

        
    #Calling on CloudConnect to connect to the tenant's Exchange Online environment via PowerShell
    Connect-ExchangeOnline -ExchangeEnvironmentName $ExchangeEnvironment

    #Connecting to MSOnline
    Connect-MsolService -AzureEnvironment $AzureEnvironment

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
    Export-UALData -ExportDir $ExportDir -UALInput $DomainData -CsvName "Domain_Operations_Export" -WorkloadType "AAD"

    #Searches for any modifications or credential modifications to an application
    Write-Verbose "Searching for 'Update application' and 'Update application ? Certificates and secrets management' in the UAL."
    $AppData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Update application","Update application ? Certificates and secrets management" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as AppUpdate_Operations_Export.csv
    Export-UALData -ExportDir $ExportDir -UALInput $AppData -CsvName "AppUpdate_Operations_Export" -WorkloadType "AAD"

    #Searches for any modifications or credential modifications to a service principal
    Write-Verbose "Searching for 'Update service principal' and 'Add service principal credentials' in the UAL."
    $SpData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Update service principal","Add service principal credentials" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as ServicePrincipal_Operations_Export.csv   
    Export-UALData -ExportDir $ExportDir -UALInput $SpData -CsvName "ServicePrincipal_Operations_Export" -WorkloadType "AAD"

    #Searches for any app role assignments to service principals, users, and groups
    Write-Verbose "Searching for 'Add app role assignment to service principal', 'Add app role assignment grant to user', and 'Add app role assignment to group' in the UAL."
    $AppRoleData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Add app role assignment" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as AppRoleAssignment_Operations_Export.csv      
    Export-UALData -ExportDir $ExportDir -UALInput $AppRoleData -CsvName "AppRoleAssignment_Operations_Export" -WorkloadType "AAD"

    #Searches for any OAuth or application consents
    Write-Verbose "Searching for 'Add OAuth2PermissionGrant' and 'Consent to application' in the UAL."
    $ConsentData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Add OAuth2PermissionGrant","Consent to application" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as Consent_Operations_Export.csv       
    Export-UALData -ExportDir $ExportDir -UALInput $ConsentData -CsvName "Consent_Operations_Export" -WorkloadType "AAD"

    #Searches for SAML token usage anomaly (UserAuthenticationValue of 16457) in the Unified Audit
    $federatedDomains = Get-MsolDomain | Where-Object {$_.Authentication -eq "Federated"}
    # Get only root domains so we can get SupportMFA status
    $rootDomains = $federatedDomains | Where-Object {$_.RootDomain -eq $null}
    # Get root domains that don't support MFA, hence Federated MFA is not expected. Note: federated MFA is still possible when SupportsMFA is false, however less likely. Check your STS configuration.
    $rootDomainsSupportMFAFalse = @()
    
    foreach ($rootDomain in $rootDomains)
    {
        $fedProps = Get-MsolDomainFederationSettings -DomainName $rootDomain.Name 
        If ($fedProps.SupportsMfa -ne $True) {
            $rootDomainsSupportMFAFalse += $rootDomain.Name
        }
    }
    # Add all child domains where its root is on the list
    $childDomainsSupportMFAFalse = @()
    $childDomains = $federatedDomains | Where-Object {$_.RootDomain -ne $null}

    foreach ($childDomain in $childDomains)
    {
        if ($childDomain.RootDomain -in $rootDomainsSupportMFAFalse){
            $childDomainsSupportMFAFalse += $childDomain.name
        }
    }

    $domainsToFlag = $rootDomainsSupportMFAFalse + $childDomainsSupportMFAFalse
    If ($null -ne $domainsToFlag)
    {
        Write-Verbose "Searching for 16457 in UserLoggedIn and UserLoginFailed operations in the UAL."
        $SAMLData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "UserLoggedIn","UserLoginFailed" -ResultSize 5000 -FreeText "16457" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
        $FilteredSAMLData = $SAMLData | Where-Object {$_.UserId.Split('@')[1] -in $domainsToFlag}
        #You can modify the resultant CSV output by changing the -CsvName parameter
        #By default, it will show up as SAMLToken_Operations_Export.csv      
        Export-UALData -ExportDir $ExportDir -UALInput $FilteredSAMLData -CsvName "SAMLToken_Operations_Export" -WorkloadType "AAD"
    } else {
        Write-Verbose "No federated domains found--16457 check will be skipped and no CSV will be produced."
    }

    #Searches for PowerShell logins into mailboxes
    Write-Verbose "Searching for PowerShell logins into mailboxes in the UAL."
    $PSMailboxData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 -Operations "MailboxLogin" -FreeText "Powershell" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as PSMailbox_Operations_Export.csv      
    Export-UALData -ExportDir $ExportDir -UALInput $PSMailboxData -CsvName "PSMailbox_Operations_Export" -WorkloadType "EXO2"

    #Searches for well-known AppID for Exchange Online PowerShell
    Write-Verbose "Searching for PowerShell logins using known PS application ids in the UAL."
    $PSLoginData1 = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000  -FreeText "a0c73c16-a7e3-4564-9a95-2bdf47383716" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #You can modify the resultant CSV output by changing the -CsvName parameter
    #By default, it will show up as PSLogin_Operations_Export.csv  
    Export-UALData -ExportDir $ExportDir -UALInput $PSLoginData1 -CsvName "PSLogin_Operations_Export" -WorkloadType "AAD"

    #Searches for well-known AppID for PowerShell
    $PSLoginData2 = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000  -FreeText "1b730954-1685-4b74-9bfd-dac224a7b894" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #The resultant CSV will be appended with the $PSLoginData* resultant CSV.
    #If you want a separate CSV with a different name, remove the -AppendType parameter (-AppendType "Append")
    #By default, it will show up as part of the PSLogin_Operations_Export.csv  
    Export-UALData -ExportDir $ExportDir -UALInput $PSLoginData2 -CsvName "PSLogin_Operations_Export" -WorkloadType "AAD" -AppendType "Append"

    #Searches for WinRM useragent string in the user logged in and user login failed operations
    $PSLoginData3 = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 -Operations "UserLoggedIn","UserLoginFailed" -FreeText "WinRM" | Select-Object -ExpandProperty AuditData | Convertfrom-Json
    #The resultant CSV will be appended with the $PSLoginData* resultant CSV.
    #If you want a separate CSV with a different name, remove the -AppendType parameter (-AppendType "Append")
    #By default, it will show up as part of the PSLogin_Operations_Export.csv 
    Export-UALData -ExportDir $ExportDir -UALInput $PSLoginData3 -CsvName "PSLogin_Operations_Export" -WorkloadType "AAD" -AppendType "Append"

    If ($AppIdInvestigation -eq "Yes"){
        If ($LicenseAnswer -eq "Yes"){
            #Searches for the AppID to see if it accessed mail items.
            Write-Verbose "Searching for $SusAppId in the MailItemsAccessed operation in the UAL."
            $SusMailItems = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "MailItemsAccessed" -ResultSize 5000 -FreeText $SusAppId -Verbose | Select-Object -ExpandProperty AuditData | Convertfrom-Json
            #You can modify the resultant CSV output by changing the -CsvName parameter
            #By default, it will show up as MailItems_Operations_Export.csv  
            Export-UALData -ExportDir $ExportDir -UALInput $SusMailItems -CsvName "MailItems_Operations_Export" -WorkloadType "EXO"
        } else {
            Write-Host "MailItemsAccessed query will be skipped as it is not present without an E5/G5 license."
        }

        #Searches for the AppID to see if it accessed SharePoint or OneDrive items
        Write-Verbose "Searching for $SusAppId in the FileAccessed and FileAccessedExtended operations in the UAL."
        $SusFileItems = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "FileAccessed","FileAccessedExtended" -ResultSize 5000 -FreeText $SusAppId -Verbose | Select-Object -ExpandProperty AuditData | Convertfrom-Json
        #You can modify the resultant CSV output by changing the -CsvName parameter
        #By default, it will show up as FileItems_Operations_Export.csv  
        Export-UALData -ExportDir $ExportDir -UALInput $SusFileItems -CsvName "FileItems_Operations_Export" -WorkloadType "SharePoint"
    }
}

Function Get-AzureDomains{

    [cmdletbinding()]Param(
        [Parameter(Mandatory=$true)]
        [string] $AzureEnvironment,
        [Parameter(Mandatory=$true)]
        [string] $ExportDir
        )

    #Connect to AzureAD
    Connect-AzureAD -AzureEnvironmentName $AzureEnvironment

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
    $DomainArr | Export-Csv $ExportDir\Domain_List.csv -NoTypeInformation
}

Function Get-AzureSPAppRoles{

    [cmdletbinding()]Param(
        [Parameter(Mandatory=$true)]
        [string] $AzureEnvironment,
        [Parameter(Mandatory=$true)]
        [string] $ExportDir
        )

    #Connect to your tenant's AzureAD environment
    Connect-AzureAD -AzureEnvironmentName $AzureEnvironment

    #Retrieve all service principals that are applications
    $SPArr = Get-AzureADServicePrincipal -All $true | Where-Object {$_.ServicePrincipalType -eq "Application"}

    #Retrieve all service principals that have a display name of Microsoft Graph
    $GraphSP = Get-AzureADServicePrincipal -All $true | Where-Object {$_.DisplayName -eq "Microsoft Graph"}

    $GraphAppRoles = $GraphSP.AppRoles | Select-Object -Property AllowedMemberTypes, Id, Value

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
    #If you want to change the default export directory, please change the $ExportDir value.
    #Otherwise, the default export is the user's home directory, Desktop folder, and ExportDir folder.
    #You can change the name of the CSV as well, the default name is "ApplicationGraphPermissions"
    $AppRolesArr | Export-Csv $ExportDir\ApplicationGraphPermissions.csv -NoTypeInformation
}

Function Export-UALData {
    Param(
        [Parameter(ValueFromPipeline=$True)]
        [Object[]]$UALInput,
        [Parameter(Mandatory=$true)]
        [String]$CsvName,
        [Parameter(Mandatory=$true)]
        [String]$WorkloadType,
        [Parameter()]
        [String]$AppendType,
        [Parameter(Mandatory=$true)]
        [string] $ExportDir
        )

        If ($UALInput.Count -eq 5000)
        {
            Write-Host 'Warning: Result set may have been truncated; narrow start/end date.'
        }

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
        } elseif ($WorkloadType -eq "SharePoint"){
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
            $DataArr | Export-csv $ExportDir\$CsvName.csv -NoTypeInformation -Append
        } Else {
            $DataArr | Export-csv $ExportDir\$CsvName.csv -NoTypeInformation
        }
        
        Remove-Variable UALInput -ErrorAction SilentlyContinue
        Remove-Variable Data -ErrorAction SilentlyContinue
        Remove-Variable DataObj -ErrorAction SilentlyContinue
        Remove-Variable DataProps -ErrorAction SilentlyContinue
        Remove-Variable DataArr -ErrorAction SilentlyContinue
}

#Function calls, if you do not need a particular check, you can comment it out below with #
Import-PSModules -ExportDir $ExportDir -Verbose
($AzureEnvironment, $ExchangeEnvironment) = Get-AzureEnvironments -AzureEnvironment $AzureEnvironment -ExchangeEnvironment $ExchangeEnvironment
If ($($ExchangeEnvironment -ne "None") -and $($NoO365 -eq $false)) {
    Get-UALData -ExportDir $ExportDir -StartDate $StartDate -EndDate $EndDate -ExchangeEnvironment $ExchangeEnvironment -AzureEnvironment $AzureEnvironment -Verbose
} 
Get-AzureDomains -AzureEnvironment $AzureEnvironment -ExportDir $ExportDir -Verbose
Get-AzureSPAppRoles -AzureEnvironment $AzureEnvironment -ExportDir $ExportDir -Verbose
New-ExcelFromCsv -ExportDir $ExportDir
