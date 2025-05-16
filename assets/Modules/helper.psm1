function GetAllADMembers
{
    [CmdletBinding()]
    [OutputType([string[]])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    $memberEmails = @()

    $groupMembers = Get-PnPAzureADGroupMember -Identity $GroupName
    if ($groupMembers)
    {
        $groupMembers | ForEach-Object {
            if ($_.UserType -eq 'Member')
            {
                $memberEmails += $_.Email
            }
            else
            {
                # nested security group
                $memberEmails += GetAllADMembers -GroupName $_.DisplayName
            }
        }
    }
    else
    {
        Write-Host "Could not find group $GroupName in M365 or AD"
    }

    Write-Output $memberEmails
}

function IsSiteCollectionAdmin
{
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]
        $UserEmail
    )

    $SiteAdminstrators = @()
    Get-PnPSiteCollectionAdmin -PipelineVariable Admin | ForEach-Object {
        if ($_.PrincipalType -eq 'SecurityGroup')
        {
            #Get Members of the Group
            $group = Get-PnPMicrosoft365Group -IncludeOwners | Where-Object { $_.Mail -eq $Admin.Email }
            if ($group)
            {
                $group.Owners | Select-Object Email | ForEach-Object {
                    $SiteAdminstrators += $_.Email
                }
                # Also include members of M365 group unless it is specifically the group owners (designated with _o)
                if (!$Admin.LoginName.EndsWith('_o'))
                {
                    $Members = Get-PnPMicrosoft365GroupMember -Identity $Group.Id -ErrorAction SilentlyContinue
                    $Members | ForEach-Object {
                        $SiteAdminstrators += $_.Email
                    }
                }
            }
            else
            {
                # Not a M365 group. Try to get the group from AD
                $SiteAdminstrators += GetAllADMembers -GroupName $Admin.Title
            }
        }
        else
        {
            $SiteAdminstrators += $Admin.Email
        }
    }

    $SiteAdminstrators = $SiteAdminstrators | Sort-Object -Unique
    $SiteAdminstrators -contains $UserEmail
}
