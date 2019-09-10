<#
    
    This is a sample Azure HTTP Trigger function to receive content links from the 
    Office 365 Management and Activity API

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

using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

#Enumerators and object to wrap the incoming request
$rawreq = @()
$rawreq = New-Object -TypeName psobject
$rawreq | Add-Member -name Content -value Content -membertype noteproperty

$requestbody = $request.RawBody | ConvertFrom-Json
$rawreq.content = $requestbody | convertto-json
#Legacy code, leaving as informational. Per documentation Next Page URI should not be present on webhook notifications

Write-Host "Next Page Header value is $($Request.Headers.NextPageUri)"

#Activity Feed webhook Body to process
[array]$contenttype = $requestBody.contenttype
#$tenantguid = $requestBody.tenantid
$clientIdIn = $requestBody.clientid
$contentId = $requestBody.contentid
[array]$contentUri = $requestBody.contentUri
$contentCreated = $requestBody.contentCreated
$contentExpiration = $requestBody.contentExpiration

Write-Host "Received Notification for ContentURI $contenturi and Content Type of $contenttype"

#Specify Client ID to match to incoming request
$ClientID = "<Customer Client ID>"
$TenantGUID = "<Customer Azure AD Tenant GUID>"

if ($clientIdIn -eq $ClientId ) {
        $contentcount = 0
        foreach($content in $contentUri)
        {
            $uri = $content + "?PublisherIdentifier=" + $TenantGUID  
            $queuehash = @{}
            $queuehash.Add("ContentURI",$uri)
            Write-Host "Adding URI $($Uri) to queue hash"
            $queuehash.Add("ContentType",$contenttype[$contentcount])
            Write-Host "Adding ContentType $($contenttype[$contentcount]) to queue hash"
            Push-OutputBinding -Name dlpqueue -Value $queuehash
            $contentcount++
            
        }
    
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 200
        Body       = "OK"
    })
}

 
