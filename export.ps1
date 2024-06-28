$ClientId = "<CLIENT ID>"
$TenantId = "<TENANT ID>"
$ClientSecret = "<CLIENT SECRET>"

$ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force

$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass

Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential

Connect-AzureAD

# Replace with the licenses you want to get
$skuPartNumbers = @(
    'ENTERPRISEPACK',
    'STANDARDPACK'
)

Import-Module ImportExcel

foreach ($skuPartNumber in $skuPartNumbers) {
    $sku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq $skuPartNumber

    $users = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($sku.SkuId))" -ConsistencyLevel eventual -All

    $userDetailsList = foreach ($user in $users) {
        $azureADUser = Get-AzureADUser -Filter "userPrincipalName eq '$($user.UserPrincipalName)'"

        if ($azureADUser) {
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                # Add fields found only on Entra
                CompanyName = $azureADUser.CompanyName
            }
        } else {
            Write-Warning "User with UserPrincipalName $($user.UserPrincipalName) not found in Azure AD."
        }
    }

    $worksheetName = "$skuPartNumber"
    $userDetailsList | Export-Excel -Path "C:\temp\userlist.xlsx" -WorksheetName $worksheetName -AutoSize -ClearSheet -BoldTopRow 
}