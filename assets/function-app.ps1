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
# searchProp - The custom search property used to tag sites
###########################################################################################################
$appUrl = "https://tenant.sharepoint.com/sites/admin";
#$azureEnv = "USGovernmentDoD";
$cert = $env:WEBSITE_LOAD_CERTIFICATES;
$clientId = $env:CLIENT_ID;
$tenant = $env:TENANT_ID;
$listName = "Site Admin Requests";
$searchProp = "BusinessUnit";
$siteAttestationDateProp = "AttestationDate";
$siteAttestationUserProp = "AttestationUser";
############################################### Global Vars ###############################################

############################################### Functions ###############################################
function GetAllADMembers {
    [CmdletBinding()]
    [OutputType([string[]])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    $memberEmails = @();
    $groupMembers = Get-PnPAzureADGroupMember -Identity $GroupName;
    if ($groupMembers) {
        $groupMembers | ForEach-Object {
            if ($_.UserType -eq 'Member') {
                $memberEmails += $_.Email;
            }
            else {
                # nested security group
                $memberEmails += GetAllADMembers -GroupName $_.DisplayName;
            }
        }
    }
    else {
        Write-Host "Could not find group $GroupName in M365 or AD";
    }

    # Return the member emails
    Write-Output $memberEmails;
}

function IsSiteCollectionAdmin {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]
        $UserEmail
    )

    $SiteAdminstrators = @();
    Get-PnPSiteCollectionAdmin -PipelineVariable Admin | ForEach-Object {
        if ($_.PrincipalType -eq 'SecurityGroup') {
            #Get Members of the Group
            $group = Get-PnPMicrosoft365Group -IncludeOwners | Where-Object { $_.Mail -eq $Admin.Email }
            if ($group) {
                $group.Owners | Select-Object Email | ForEach-Object {
                    $SiteAdminstrators += $_.Email;
                }
                # Also include members of M365 group unless it is specifically the group owners (designated with _o)
                if (!$Admin.LoginName.EndsWith('_o')) {
                    $Members = Get-PnPMicrosoft365GroupMember -Identity $Group.Id -ErrorAction SilentlyContinue;
                    $Members | ForEach-Object {
                        $SiteAdminstrators += $_.Email;
                    }
                }
            }
            else {
                # Not a M365 group. Try to get the group from AD
                $SiteAdminstrators += GetAllADMembers -GroupName $Admin.Title;
            }
        }
        else {
            $SiteAdminstrators += $Admin.Email;
        }
    }

    # Sort and remove duplicates
    $SiteAdminstrators = $SiteAdminstrators | Sort-Object -Unique;

    # Return true/false if the user email is in the site admin array
    $SiteAdminstrators -contains $UserEmail;
}
############################################### Functions ###############################################

############################################### Validation ###############################################
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
############################################### Validation ###############################################

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
        $requestType = $item["RequestType"];

        # Log
        Write-Host "Processing site $siteUrl";
        Write-Host "Request Type: $requestType";
        Write-Host "Request Value: $value";

        # Connect to the site
        Connect-PnPOnline -Url $siteUrl -Tenant $tenant -ClientId $clientId -Thumbprint $cert;
        #Connect-PnPOnline -Url $siteUrl -Tenant $tenant -ClientId $clientId -Thumbprint $cert -AzureEnvironment $azureEnv;

        # Check to make sure the user requesting the change is an admin
        if (-not(IsSiteCollectionAdmin -UserEmail $item["Author"].EMail)) {
            # Log
            Write-Host "The user is not a site collection admin";

            # Set the output
            $output = "The user is not a site collection admin.";

            # Set the status code
            $statusCode = 401;

            # Disconnect from the site
            Disconnect-PnPOnline;

            # Return
            return;
        }

        # Process the setting, based on the request type
        switch ($requestType) {
            "App Catalog" {
                # See if we are enabling the app catalog
                if ($value -eq "true") {
                    # Log
                    Write-Host "Enabling the app catalog...";

                    # Add the site collection app catalog
                    Add-PnPSiteCollectionAppCatalog -site $siteUrl;

                    # Log
                    Write-Host "Waiting 10 seconds for the client side assets library to be available...";

                    # Wait for the library to be created
                    Start-Sleep -Seconds 10;

                    # Log
                    Write-Host "Setting exemption for download...";

                    # Add Client Site Asset library exception
                    Set-PnPList -Identity "Client Side Assets" -ExemptFromBlockDownloadOfNonViewableFiles $true;

                    # Set the output
                    $output = "The app catalog has been enabled for this site collection.";
                }
                else {
                    # Log
                    Write-Host "Disabling the app catalog...";

                    # Remove the site collection app catalog
                    Remove-PnPSiteCollectionAppCatalog -site $siteUrl;

                    # Log
                    Write-Host "Removing exemption for download...";

                    # Remove Client Site Asset library exception
                    Set-PnPList -Identity "Client Side Assets" -ExemptFromBlockDownloadOfNonViewableFiles $false;

                    # Set the output
                    $output = "The app catalog has been disabled for this site collection.";
                }
            }
            "Client Side Assets" {
                # See if we are enabling the app catalog
                if ($value -eq "true") {
                    # Log
                    Write-Host "Setting exemption for download...";

                    # Add Client Site Asset library exception
                    Set-PnPList -Identity "Client Side Assets" -ExemptFromBlockDownloadOfNonViewableFiles $true;

                    # Set the output
                    $output = "The client side assets library will bypass the block download policy.";
                }
                else {
                    # Log
                    Write-Host "Setting exemption for download...";

                    # Add Client Site Asset library exception
                    Set-PnPList -Identity "Client Side Assets" -ExemptFromBlockDownloadOfNonViewableFiles $false;

                    # Set the output
                    $output = "The client side assets library will not bypass the block download policy.";
                }
            }
            "Custom Script" {
                # See if we are enabling custom scripts
                if ($value -eq "true") {
                    # Log
                    Write-Host "Setting NoScriptSite property to false...";

                    # Enable custom scripts
                    #Set-PnPTenantSite -Identity $siteUrl -DenyAddAndCustomizePages $false;
                    Set-PnPSite -Identity $siteUrl -NoScriptSite $false;

                    # Set the output
                    $output = "Custom Script feature has been enabled on your site. Please note that the setting will revert after 24 hours.";
                }
                else {
                    # Log
                    Write-Host "Setting NoScriptSite property to true...";

                    # Disable custom scripts
                    #Set-PnPTenantSite -Identity $siteUrl -DenyAddAndCustomizePages $true;
                    Set-PnPSite -Identity $siteUrl -NoScriptSite $true;

                    # Set the output
                    $output = "Custom Script feature has been disabled on your site. Please note that the setting will revert after 24 hours.";
                }
            }
            "Custom Search Property" {
                # Log
                Write-Host "Setting NoScriptSite to false...";

                # Enable custom scripts
                #Set-PnPTenantSite -Identity $siteUrl -DenyAddAndCustomizePages $true;
                Set-PnPSite -Identity $siteUrl -NoScriptSite $false;

                # See if a value exists
                if ([string]::IsNullOrEmpty($value)) {
                    # Log
                    Write-Host "Removing the key $searchProp from the property bag...";

                    # Remove the site property
                    Remove-PnPPropertyBagValue -Key $searchProp -Force;

                    # Set the output
                    $output = "The property bag $searchProp value was removed.";
                }
                else {
                    # Log
                    Write-Host "Adding the key $searchProp to the property bag with value $value...";

                    # Update the site property
                    Set-PnPPropertyBagValue -Key $searchProp -Value $value -Indexed;

                    # Set the output
                    $output = "The property bag $searchProp value was set to $value.";
                }

                # Log
                Write-Host "Setting NoScriptSite to true...";

                # Disable custom scripts
                #Set-PnPTenantSite -Identity $siteUrl -DenyAddAndCustomizePages $true;
                Set-PnPSite -Identity $siteUrl -NoScriptSite $true;
            }
            "Company Wide Sharing Links" {
                # See if we are disabling company wide sharing links
                if ($value -eq "true") {
                    # Log
                    Write-Host "Setting the DisableCompanyWideSharingLinks to true...";

                    # Disable the company wide sharing links
                    Set-PnPSite -Identity $siteUrl -DisableCompanyWideSharingLinks $true;

                    # Set the output
                    $output = "The company wide sharing links has been disabled for this site collection.";
                }
                else {
                    # Log
                    Write-Host "Setting the DisableCompanyWideSharingLinks to false...";

                    # Enable the company wide sharing links
                    Set-PnPSite -Identity $siteUrl -DisableCompanyWideSharingLinks $false;

                    # Set the output
                    $output = "The company wide sharing links has been enabled for this site collection.";
                }
            }
            "Increase Storage" {
                # See if we are enabling custom scripts
                if ($value -eq "true") {
                    # Log
                    Write-Host "Increasing the storage size to 25TB...";

                    # Set the max and warning value (80%)
                    # 1GB  = 1024 * 1024      = 1048576  (838860)
                    # 5GB  = 5 * 1024 * 1024  = 5242880  (4194304)
                    # 10GB = 10 * 1024 * 1024 = 10485760 (8388608)
                    # 15GB = 15 * 1024 * 1024 = 15728640 (12582912)
                    # 20GB = 20 * 1024 * 1024 = 20971520 (16777216)
                    # 25GB = 25 * 1024 * 1024 = 26214400 (20971520)
                    Set-PnPSite -Identity $siteUrl -StorageMaximumLevel 26214400 -StorageWarningLevel 20971520;

                    # Set the output
                    $output = "The site storage size has increased to 25TB.";
                }
            }
            "Lock State" {
                # Log
                Write-Host "Setting LockState to $value...";

                # Set the lock state
                Set-PnPSite -Identity $siteUrl -LockState $value;

                # Set the output
                $output = "The lock state has been set to '$value' for this site collection.";
            }
            "No Crawl" {
                # Log
                Write-Host "Enabling custom scripts...";

                # Enable custom scripts
                Set-PnPSite -Identity $siteUrl -NoScriptSite $false;

                # Log
                Write-Host "Setting NoCrawl to $value...";

                # See if we are setting turning it off
                if($value -eq "true") {
                    # Set the no crawl property
                    Set-PnPWeb -NoCrawl:$true;
                } else {
                    # Set the no crawl property
                    Set-PnPWeb -NoCrawl:$false;
                }

                # Log
                Write-Host "Disabling custom scripts...";

                # Disable custom scripts
                Set-PnPSite -Identity $siteUrl -NoScriptSite $true;
            }
            "Site Attestation" {
                # Log
                Write-Host "Setting NoScriptSite to false...";

                # Enable custom scripts
                Set-PnPSite -Identity $siteUrl -NoScriptSite $false;

                # Get the current date/time in ISO 8601 format
                $currentDateTime = Get-Date -Format "o";

                # Log
                Write-Host "Adding the key $siteAttestationDateProp to the property bag with value $currentDateTime...";
                Write-Host "Adding the key $siteAttestationUserProp to the property bag with value $value...";

                # Update the site property
                Set-PnPPropertyBagValue -Key $siteAttestationDateProp -Value $currentDateTime -Indexed;
                Set-PnPPropertyBagValue -Key $siteAttestationUserProp -Value $value -Indexed;

                # Set the output
                $output = "The site attestation was completed for $value.";

                # Log
                Write-Host "Setting NoScriptSite to true...";

                # Disable custom scripts
                Set-PnPSite -Identity $siteUrl -NoScriptSite $true;
            }
        }

        # Log
        Write-Host "Setting the request item status to Completed...";

        # Connect to the main site
        Connect-PnPOnline -Url $appUrl -Tenant $tenant -ClientId $clientId -Thumbprint $cert;
        #Connect-PnPOnline -Url $appUrl -Tenant $tenant -ClientId $clientId -Thumbprint $cert -AzureEnvironment $azureEnv;

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
