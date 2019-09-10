# Office365ActivityFunctions

Azure functions for interacting with the Office 365 Management and Activity API

# Overview

This sample utilizes an Azure Function app with two funcitons to consume the Office 365 Management and Activity API events. The first function is an HTTP Trigger Function designed to be the webhook endpoint for the Managment and Activity API. This Function retrieves Content links from the API, and passes them to an Azure Storage Queue.

The Second function is a Queue Trigger function which triggers on queues messages added by the first function. This function utilizes a Client ID and secret associated with an Azure AD registered application with permissions to the Management and Activity API. The secret is stored in Azure Key Vault and accessed only at runtime. The function process each content link as read from the queue, and passes on the JSON response to an Azure Storage Table. The Azure Storage Table binding as part of Azure Functions handles the de-serialization of the JSON response.



