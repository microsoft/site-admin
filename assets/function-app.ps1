using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

############################################### Global Vars ###############################################
# The variables need to be filled out for the PnP Auth connection to work
# appUrl - The site containing the app and list
# azureEnv - The environment the tenant belongs to
# cert - The thumbprint of the certificate
# clientId - The client id of the app registration
# tenant - The domain of the tenant
# listName - The list name containing the requests
###########################################################################################################
$appUrl = "https://tenant.sharepoint.com/sites/admin";
#$azureEnv = "USGovernmentDoD";
$cert = $env:WEBSITE_LOAD_CERTIFICATES;
$clientId = $env:CLIENT_ID;
#$clientSecret = $env:CLIENT_SECRET
$tenant = $env:TENANT_ID;
$listName = "Site Admin Requests";
############################################### Global Vars ###############################################

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function triggered..."

# Get the request id from the body
$itemId = $Request.Body.requestId;

# Log
Write-Host "The item id provided was: $itemId";

# Validate the item id
if (-not($itemId -gt 0)) {
    # Respond w/ an error
    Write-Host "Item Id failed validation";
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = 401
            Body       = "The request id is invalid."
        });

    return;
}

# Log
Write-Host "Item Id was validated";

############################################### SP Connection ###############################################
Write-Host "App Url: $appUrl";
Write-Host "Azure Environment: $azureEnv";
Write-Host "Client ID: $clientId";
Write-Host "Tenant: $tenant";
Write-Host "List Name: $listName";
Write-Host "Connecting with PnP PowerShell..."
Connect-PnPOnline -Url $appUrl -Tenant $tenant -ClientId $clientId -Thumbprint $cert;
#Connect-PnPOnline -Url $appUrl -Tenant $tenant -ClientId $clientId -Thumbprint $cert -AzureEnvironment $azureEnv;
############################################### SP Connection ###############################################

############################################### Main App ###############################################
$output = "";
$statusCode = [HttpStatusCode]::OK;

# Get the list items that need attention
$item = Get-PnPListItem -Id $itemId -List $listName;
$value = $item["RequestValue"];
if ($item -ne $null) {
    $siteUrl = $item["Title"];

    # Ensure the status is "New" or "Error"
    $currentStatus = $item["Status"];
    if ($currentStatus -eq "Completed") {
        # The item is already processed
        $output = "The item has already been processed.";
    }
    else {
        # Host
        Write-Host "Processing site $siteUrl";

        # Process the setting, based on the request type
        switch ($item["RequestType"]) {
            "App Catalog" {
                # See if we are enabling the app catalog
                if ($value -eq "true") {
                    # Add the site collection app catalog
                    Add-PnPSiteCollectionAppCatalog -site $siteUrl;

                    # Add Client Site Asset library exception
                    Set-PnPList -Identity "Client Side Assets" -ExemptFromBlockDownloadOfNonViewableFiles $true;

                    # Set the output
                    $output = "The app catalog has been enabled for this site collection.";
                }
                else {
                    # Remove the site collection app catalog
                    Remove-PnPSiteCollectionAppCatalog -site $siteUrl;

                    # Remove Client Site Asset library exception
                    Set-PnPList -Identity "Client Side Assets" -ExemptFromBlockDownloadOfNonViewableFiles $false;

                    # Set the output
                    $output = "The app catalog has been disabled for this site collection.";
                }
            }
            "Custom Script" {
                # See if we are enabling custom scripts
                if ($value -eq "true") {
                    # Enable custom scripts
                    Set-PnPSite -Identity $siteUrl -NoScriptSite $false;

                    # Set the output
                    $output = "Custom Script feature has been enabled on your site. Please note that the setting will revert after 24 hours.";
                }
                else {
                    # Disable custom scripts
                    Set-PnPSite -Identity $siteUrl -NoScriptSite $true;

                    # Set the output
                    $output = "Custom Script feature has been disabled on your site. Please note that the setting will revert after 24 hours.";
                }
            }
            "Company Wide Sharing Links" {
                # See if we are disabling company wide sharing links
                if ($value -eq "true") {
                    # Disable the company wide sharing links
                    Set-PnPSite -Identity $siteUrl -DisableCompanyWideSharingLinks $true;

                    # Set the output
                    $output = "The company wide sharing links has been disabled for this site collection.";
                }
                else {
                    # Enable the company wide sharing links
                    Set-PnPSite -Identity $siteUrl -DisableCompanyWideSharingLinks $true;

                    # Set the output
                    $output = "The company wide sharing links has been enabled for this site collection.";
                }
            }
            "Increase Storage" {
                # See if we are enabling custom scripts
                if ($value -eq "true") {
                    # Set the max to 25TB
                    Set-PnPSite -Identity $siteUrl -StorageMaximumLevel 26214400 -StorageWarningLevel 24136192;
                }
            }
            "Lock State" {
                # Set the lock state
                Set-PnPSite -Identity $siteUrl -LockState $value;

                # Set the output
                $output = "The lock state has been set to '$value' for this site collection.";
            }
        }

        # Update the item status
        Set-PnpListItem -List $listName -Identity $item.Id -Values @{ "Status" = "Completed" };

        # Host
        Write-Host "Completed...";
    }
}
else {
    # Log
    Write-Host "Unable to find the item by id: $itemId";

    # Set the output
    $output = "Unable to find the item by id: $itemId";

    # Set the status code
    $statusCode = 503;
}

############################################### Main App ###############################################

############################################### Disconnect ###############################################
Disconnect-PnPOnline;
############################################### Disconnect ###############################################

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $statusCode
    Body       = $output
});
