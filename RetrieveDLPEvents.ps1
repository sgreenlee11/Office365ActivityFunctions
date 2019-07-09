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

Write-Host "Content URI $($QueueItem.ContentURI)"
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
$contentreq = Invoke-WebRequest -Method Get -Headers $headerParams -Uri $uri -ContentType application/json

Push-OutputBinding -Name outputDLPEventHubMessage -Value $contentreq.Content