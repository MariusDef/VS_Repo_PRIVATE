Get-DistributionGroup | Where-Object { $_.ManagedBy } | ForEach-Object {
  $managedBy = $_.ManagedBy
  foreach ( $managerId in $managedBy ) {
    New-Object PSObject -Property @{
      "Group" = $_.Name
      "Manager" = $managerId.Name
    } | Select-Object Group,Manager
  }
}