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
        If ($_.PrincipalType -eq 'SecurityGroup')
        {
            #Get Members of the Group
            $Group = Get-PnPMicrosoft365Group -IncludeOwners | Where-Object { $_.Mail -eq $Admin.Email }
            $Group.Owners | Select-Object Email | ForEach-Object {
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
        Else
        {
            $SiteAdminstrators += $Admin.Email
        }
    }

    $SiteAdminstrators = $SiteAdminstrators | Sort-Object -Unique
    $SiteAdminstrators -contains $UserEmail
}
