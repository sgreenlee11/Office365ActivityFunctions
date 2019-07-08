<#
    
    This configuration file provides the settings to be used with the FullExchangeInstallConfig_prod.psd1 DSC configuration.

    The sample scripts are not supported under any Microsoft standard support 
    program or service. The sample scripts are provided AS IS without warranty  
    of any kind. Microsoft further disclaims all implied warranties including,  
    without limitation, any implied warranties of merchantability or of fitness for 
    a particular purpose. The entire risk arising out of the use or performance of  
    the sample scripts and documentation remains with you. In no event shall 
    Microsoft, its authors, or anyone else involved in the creation, production, or 
    delivery of the scripts be liable for any damages whatsoever (including, 
    without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use 
    of or inability to use the sample scripts or documentation, even if Microsoft 
    has been advised of the possibility of such damages.
#>
# Input bindings are passed in via param block.
param([System.Collections.Hashtable] $QueueItem, $TriggerMetadata)

# Write out the queue message and insertion time to the information log.
Write-Host "PowerShell queue trigger function processed work item: $QueueItem"
Write-Host "Queue item insertion time: $($TriggerMetadata.InsertionTime)"

Write-Host "Conent URI $($QueueItem.ContentURI)"
Write-Host "Content Type $($QueueItem.Contenttype)"

#Sign in Parameters - Update with customer specific settings
$ClientID = "<Customer Client ID>"
$loginURL = "https://login.windows.net"
$tenantdomain = "<Customer Onmicrosoft.com Domain>"
$TenantGUID = "<Customer Azure AD Tenant GUID"
$resource = "https://manage.office.com"

#Obtain token for key vault
# Our Key Vault Credential that we want to retreive URI - Update with customer
$vaultSecretURI = "<Customer Azure Vault Secret URL"

#Values for local token service

$apiVersion = "2017-09-01"
$resourceURI = "https://vault.azure.net"
$tokenAuthURI = $env:MSI_ENDPOINT + "?resource=$resourceURI&api-version=$apiVersion"
$tokenResponse = Invoke-RestMethod -Method Get -Headers @{"Secret"="$env:MSI_SECRET"} -Uri $tokenAuthURI

# Use Key Vault AuthN Token to create Request Header
$requestHeader = @{ Authorization = "Bearer $($tokenresponse.access_token)" }
# Call the Vault and Retrieve Creds
$secret = Invoke-RestMethod -Method GET -Uri $vaultSecretURI -ContentType 'application/json' -Headers $requestHeader

# Get an Oauth 2 access token based on client id, secret and tenant domain
$body = @{grant_type = "client_credentials"; resource = $resource; client_id = $ClientID; client_secret = $secret.value }

#Request Management and Activity API token and create header

$oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body 
$headerParams = @{'Authorization' = "$($oauth.token_type) $($oauth.access_token)" }

#Make Content Request
Write-Host "Retrieving API Content using Key Vault Credential"
$uri = $QueueItem.ContentURI
Write-Host "Content URI is $($Uri)"
$restrequest = Invoke-RestMethod -Headers $headerParams -Uri $uri
Write-Host "Retrieving $($restrequest.Count) Events for source $($ContentType)"
$dlpevents = New-Object System.Collections.ArrayList
foreach ($req in $restrequest) {
    $dlpevent = New-Object psobject
    $dlpevent | Add-Member -MemberType NoteProperty -Name "CreationTime" -Value $req.CreationTime
    $dlpevent | Add-Member -MemberType NoteProperty -Name "Id" -Value $req.id
    $dlpevent | Add-Member -MemberType NoteProperty -Name "Operation" -Value $req.Operation
    $dlpevent | Add-Member -MemberType NoteProperty -Name "OrganizationId" -Value $req.OrganizationId
    $dlpevent | Add-Member -MemberType NoteProperty -Name "RecordType" -Value $req.RecordType
    $dlpevent | Add-Member -MemberType NoteProperty -Name "UserKey" -Value $req.UserKey
    $dlpevent | Add-Member -MemberType NoteProperty -Name "UserType" -Value $req.UserType
    $dlpevent | Add-Member -MemberType NoteProperty -Name "Version" -Value $req.Version
    $dlpevent | Add-Member -MemberType NoteProperty -Name "WorkLoad" -Value $req.WorkLoad
    $dlpevent | Add-Member -MemberType NoteProperty -Name "ObjectId" -Value $req.ObjectId
    $dlpevent | Add-Member -MemberType NoteProperty -Name "UserId" -Value $req.UserId
    $dlpevent | Add-Member -MemberType NoteProperty -Name "IncidentId" -Value $req.IncidentId
    $dlpevent | Add-Member -MemberType NoteProperty -Name "SensitiveInfoDetctionIsIncluded" -Value $req.SensitiveInfoDectionIsIncluded

    foreach ($prop in ($req.ExchangeMetaData | Get-Member | ? { $_.membertype -Match "Property" })) {
        $dlpevent | Add-Member -MemberType NoteProperty -Name $prop.name -Value $req.ExchangeMetaData.$($prop.name)
    }
    $dlpevent | Add-Member -MemberType NoteProperty -Name "PolicyId" -Value $req.PolicyDetails.PolicyId
    $dlpevent | Add-Member -MemberType NoteProperty -Name "PolicyName" -Value $req.PolicyDetails.PolicyName
    [void]$dlpevents.Add($dlpevent)

<# Reference code for pushing to Azure Table Storage
    $tablehash.Add("RowKey", ([GUID]::NewGuid()))
    $tablehash.Add("PartitionKey", $QueueItem.ContentType)
    $tablehash.Add("CreationTime", $req.CreationTime)
    $tablehash.Add("Id", $req.Id)
    $tablehash.Add("Operation", $req.Operation)
    $tablehash.Add("OrganizationId", $req.OrganizationId)
    $tablehash.Add("RecordType", $req.RecordType)
    $tablehash.Add("UserKey", $req.UserKey)
    $tablehash.Add("UserType", $req.UserType)
    $tablehash.Add("Version", $req.Version)
    $tablehash.Add("Workload", $req.Workload)
    $tablehash.Add("ObjectId", $req.ObjectId)
    $tablehash.Add("UserId", $req.UserId)
    $tablehash.Add("IncidentId", $req.IncidentId)
    $tablehash.Add("SensitiveInfoDectionIsIncluded", $req.SensitiveInfoDectionIsIncluded)
    foreach ($prop in ($req.ExchangeMetaData | Get-Member | ? { $_.membertype -Match "Property" })) {
        $tablehash.add($prop.name, $req.ExchangeMetaData.$($prop.name))
    }
    $tablehash.Add("PolicyId", $req.PolicyDetails.PolicyId)
    $tablehash.Add("PolicyName", $req.PolicyDetails.PolicyName)
    Push-OutputBinding -Name O365AuditData -Value $tablehash
#>
}


Push-OutputBinding -Name dlpblob -Value ($dlpevents | ConvertTo-Csv)
