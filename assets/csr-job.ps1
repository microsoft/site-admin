############################################### Global Vars ###############################################
# The variables need to be filled out for the PnP Auth connection to work
# appUrl - The site containing the app and list
# cert - The thumbprint of the certificate
# clientId - The client id of the app registration
# tenant - The domain of the tenant
# listName - The list name containing the requests
###########################################################################################################
$appUrl = "";
$cert = "";
$clientId = "";
$tenant = "";
$listName = "Custom Script Requests";
############################################### Global Vars ###############################################

############################################### Library Check ###############################################
# Validate PSVersion
Write-Host "Validating PSVersion 7.2+";
if ($PSVersionTable.PSVersion.Major -lt 7 -or ($PSVersionTable.PSVersion.Major -eq 7 -and $PSVersionTable.PSVersion.Minor -lt 2)) {
    Write-Host "Version of PowerShell must be 7.2 or greater.";
    return;
}

# Import the PnP Library
Write-Host "Validating PnP Library";
$lib = Get-InstalledModule -Name "PnP.PowerShell";
if ($null -eq $lib) {
    Install-Module -Name "PnP.PowerShell" -Scope CurrentUser -Force;
}
############################################### Library Check ###############################################

############################################### SP Connection ###############################################
Write-Host "Tenant: $tenant";
Write-Host "App Url: $appUrl";
Write-Host "List Name: $listName";
Write-Host "Connecting with PnP PowerShell..."
Connect-PnPOnline -Url $appUrl -Tenant $tenant -ClientId $clientId -Thumbprint $cert;
############################################### SP Connection ###############################################


############################################### Main App ###############################################
# Get the list items that need attention
$query = '<View>
    <ViewFields>
        <FieldRef Name="Id" />
        <FieldRef Name="Title" />
        <FieldRef Name="Status" />
    </ViewFields>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name="Status" />
                <Value Type="Text">New</Value>
            </Eq>
        </Where>
    </Query>
</View>';
$items = Get-PnPListItem -List $listName -Query $query -PageSize 1000 -ScriptBlock {
    Param($items)
    $items.Context.ExecuteQuery()
} | foreach-object {
    $item = $_;
    $siteUrl = $item["Title"];

    # Host
    Write-Host "Processing site $siteUrl";

    # Enable the setting
    Set-PnPSite -Identity $siteUrl -NoScriptSite $false;

    # Update the item status
    Set-PnpListItem -List $listName -Identity $item.Id -Values @{ "Status" = "Processed" };

    # Host
    Write-Host "Completed...";
};

############################################### Main App ###############################################

############################################### Disconnect ###############################################
Disconnect-PnPOnline;
############################################### Disconnect ###############################################