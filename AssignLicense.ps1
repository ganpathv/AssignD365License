[CmdletBinding()]
Param
(
    [Parameter(ParameterSetName="ExportCsv")][switch]$ExportCsv,
    [Parameter(ParameterSetName="ExportCsv", Mandatory=$true)][string]$OutputFile,
    [Parameter(ParameterSetName="AssignLicenseFromO365")][switch]$AssignFromO365,
	[Parameter(ParameterSetName="AssignLicenseFromCsv")][switch]$AssignFromCsv,
    [Parameter(ParameterSetName="AssignLicenseFromCsv", Mandatory=$true)][string]$InputFile,
	[switch]$RemoveCrmStandardLicense
)

$CrmOnlineProfessionalPlan = "CRMSTANDARD"
$Dyn365EnterprisePlan = "DYN365_ENTERPRISE_PLAN1"
[int]$userCount = 0;

Import-Module MSOnline

function ConnectToMsolService()
{
    try
    {
        $credential = Get-Credential

        Connect-MsolService -Credential $credential -ErrorAction Stop -ErrorVariable $errorConnecting
        
        if($errorConnecting)
        {
            Write-Host "Error connecting to O365 service.\n" $errorConnecting
            return $false
        }
        else
        {
            Write-Host "Connected to O365 Service"

            return $true
        }
    }
    catch
    {
        Write-Host "Error connecting to O365 Service"
        return $false
    }
}

function LoadUsers()
{
    try
    {
        if($PSCmdlet.ParameterSetName -eq "AssignLicenseFromCsv")
        {
            If(Test-path $InputFile)
            {
                $enabledUsers = Import-Csv $InputFile
                return $enabledUsers
            }
        }
        elseif($PSCmdlet.ParameterSetName -eq "AssignLicenseFromO365")
        {
            # Get all Enabled and licensed users
            $enabledUsers = Get-MsolUser -EnabledFilter EnabledOnly | where {$_.isLicensed -eq $true}
            return $enabledUsers
        }
    }
    catch
    {
        Write-Host "Error loading users"
        return $false
    }
}

function AssignLicenses()
{
    # Get all Enabled and licensed users
    $enabledUsers = LoadUsers
    
    # Iterate through the list of users retrieved
    Write-Host "Found " $enabledUsers.Count " enabled and licensed users to process"

	$LicenseToRemove = Get-MsolAccountSku | select AccountSkuId | where {$_ -match $CrmOnlineProfessionalPlan}
	
	# All Plans are prefixed with Domain name and a :. e.g. Contoso:CRMSTANDARD
	$skuPrefix = $LicenseToRemove.AccountSkuId.Substring(0,$LicenseToRemove.AccountSkuId.IndexOf((":"))+1)

	$LicenseToAdd = $skuPrefix + $Dyn365EnterprisePlan

    foreach($enabledUser in $enabledUsers)
    {
        $UPN = $enabledUser.UserPrincipalName

        # Get the list of licenses
        $userLicenses = Get-MsolUser -UserPrincipalName $UPN 

        $license = $userLicenses.Licenses | select AccountSkuId | where {$_ -match $LicenseToRemove}

        if($license)
        {
            Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $LicenseToAdd -ErrorAction SilentlyContinue -ErrorVariable $assignError

            if($assignError)
            {
                Write-Host "An error occurred while assigning the license to for user" $enabledUser.DisplayName "("$UPN"). Existing license will not be removed for the user."
            }
            else
            {
				if($RemoveLicenses -eq $true)
				{
					Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $LicenseToRemove  -ErrorAction SilentlyContinue $removeError
                    if($removeError)
                    {
                        Write-Host "An error occurred while removing the license to for user" $enabledUser.DisplayName "("$UPN")."
                    }
                    else
                    {
                        Write-Host "Removed $CrmOnlineProfessionalPlan license plan and assigned $Dyn365EnterprisePlan for user:" $enabledUser.DisplayName "("$UPN")"
                    }
				}
                else
                {
                    Write-Host "Assigned $Dyn365EnterprisePlan for user:" $enabledUser.DisplayName "("$UPN")"
                }
            }
            $userCount++
        }
    }
    Write-Host "Reassigned the new $Dyn365EnterprisePlan for $userCount users"
}

function ExportUsersToCsv($outFile)
{
    try
    {
        if($outFile.IndexOf(".csv") -gt 0)
        {
            # Get all Enabled and licensed users
            $enabledUsers = Get-MsolUser -EnabledFilter All #| where {$_.isLicensed -eq $true}
    
            # Iterate through the list of users retrieved
            $enabledUsers | select DisplayName, UserPrincipalName, IsLicensed, Usagelocation | Export-Csv $outFile

            Write-Host "User list written to $outFile"
        }
        else
        {
            Write-Host "Please provide full path to the csv file"
        }
    }
    catch [System.Exception]
    {
        Write-Error $_.Exception.ToString()
        exit -1
    }
}

try
{
    if(ConnectToMsolService)
    {
        switch($PSCmdlet.ParameterSetName)
        {
            "ExportCsv"
            {
                ExportUsersToCsv $OutputFile
            }
            "AssignLicenseFromO365"
            {
                # Though the same method is called, segregating the call just to provide the ability to extend this if required
                AssignLicenses
            }
            "AssignLicenseFromCsv"
            {
                AssignLicenses
            }
        }
    }

    
}
catch [System.Exception]
{
    Write-Error $_.Exception.ToString()
    exit -1
}
    
