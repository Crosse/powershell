param (
        [switch]
        $RefreshDatastoreUsage=$false,

        [switch]
        $VMUsage=$true,

        [switch]
        $DatastoreUsage=$true,

        [ValidateScript({ if ($VMUsage) { Resolve-Path ([System.IO.FileInfo]$_).Directory }})]
        [string]
        $VMUsageFilePath="vm_datastore_usage.csv",

        [ValidateScript({ if ($DatastoreUsage) { Resolve-Path ([System.IO.FileInfo]$_).Directory }})]
        [string]
        $DatastoreUsageFilePath="datastore_usage.csv"
      )

if ($global:DefaultVIServer -eq $null) {
    Connect-VIServer -Menu
}

$vmfs = Get-Datastore 
$allvms = Get-VM

if ($RefreshDatastoreUsage)
{
    $count = $vmfs.Count
    $i = 0
    foreach ($lun in $vmfs) {
        Write-Progress  -Activity "Refreshing Datastore Storage Information" `
        -Status "$lun" -PercentComplete ($i / $count * 100)
        ($lun | Get-View).RefreshDatastoreStorageInfo()
        $i++
    }
}
Write-Progress -Activity "Refreshing Datastore Storage Information" -Status "Completed" -Completed:$true

if ($VMUsage) {
    $count = $allvms.Count
    $i = 0
    $report = @()                                         # blank data structure
    foreach ($vm in $allvms) { 
        Write-Progress  -Activity "Processing VM usage information" `
        -Status "$vm" -PercentComplete ($i / $count * 100)

        $vmview = $vm | Get-View                    # get-view to see details of the FM
        foreach($disk in $vmview.Storage.PerDatastoreUsage) {      # Disks for the VM
            $dsview = (Get-View $disk.Datastore)                    # Datastore used by the disks

            $row = "" | select VMName, Datastore, VMSize_MB, VMUsed_MB, Percent     # blank row
            $row.VMName = $vmview.Config.Name                         # Add the data to the row
            $row.Datastore = $dsview.Name
            $row.VMSize_MB = (($disk.Committed+$disk.Uncommitted)/1024/1024)
            $row.VMUsed_MB = (($disk.Committed)/1024/1024)
            $row.Percent = [int](($row.VMUsed_MB / $row.VMSize_MB)*100)
            $report += $row                                                                  # Add the row to the structure
        } 
        $i++
    } 
    Write-Progress -Activity "Processing datastore usage information" -Status "Completed" -Completed:$true

    $report | Export-Csv $VMUsageFilePath -NoTypeInformation     # dump the report to .csv
}

if ($DatastoreUsage) {
    $count = $vmfs.Count
    $i = 0
    $DSReport = @()                         # blank data structure for datastore information
    foreach ($lun in $vmfs) {          
        Write-Progress  -Activity "Processing datastore usage information" `
        -Status "$lun" -PercentComplete ($i / $count * 100)
        $VMSizeSum = 0                    # We will sum the data from the previous report for this LUN

        foreach ($row in $report) {     # Generate sum for this LUN
            if ($row.Datastore -eq $lun.Name) {$VMSizeSum += $row.VMSIZE_MB}
        }
# Create a blank row and add data to it.
        $DSRow = "" | select Datastore_Name,Capacity_MB,  FreeSpace_MB, Allocated_MB, Unallocated_MB
        $DSRow.Datastore_Name = $lun.Name
        $DSRow.Capacity_MB = $lun.CapacityMB
        $DSRow.FreeSpace_MB = $lun.FreeSpaceMB
        $DSRow.Allocated_MB = [int]$VMSizeSum
        $DSRow.Unallocated_MB = $lun.CapacityMB - [int]$VMSizeSum     # NB that if we have overallocated disk                                                  # space this will be a negative number
        $DSReport += $DSRow               # add the row to the structure.
        $i++
    }     
    Write-Progress -Activity "Processing datastore usage information" -Status "Completed" -Completed:$true

    $DSReport | Export-Csv $DatastoreUsageFilePath -NoTypeInformation     # dump report to .csv
}
