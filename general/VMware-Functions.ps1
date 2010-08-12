if (Test-Path function:SVMotion-VM) { Remove-Item function:SVMotion-VM }
Function global:SVMotion-VM {
	Param(
		[VMware.VimAutomation.Client20.VirtualMachineImpl[]]
		$vms,

		[VMware.VimAutomation.Client20.DatastoreImpl]
		$destination
	)

	$datastoreView = get-view $destination.ID
	$relocationSpec = new-object VMware.Vim.VirtualMachineRelocateSpec
	$relocationSpec.Datastore = $datastoreView.MoRef

	$tasks = @()
	foreach ($vm in $vms) {
		$vmView = get-view $vm.ID
		$tasks += $vmView.RelocateVM_Task($relocationSpec)
	}

	# In 1.0 GA we will have a better way to deal with tasks created in this way. Until
	# then we just let them run with no results returned.
}

Write-Host "`tAdded SVMotion-VM to global functions." -Fore White

if (Test-Path function:Get-DatastoreFiles) { Remove-Item function:Get-DatastoreFiles }
Function global:Get-DatastoreFiles {
	Param(
		[VMware.VimAutomation.Client20.DatastoreImpl[]]
		$datastores
	)

	$ret = @()
	foreach ($datastore in $datastores) {
		$datastoreView = get-view $datastore.id
		$datastoreBrowser = get-view $datastoreView.browser
		$name = $datastore.name
		$spec = new-object vmware.vim.HostDatastoreBrowserSearchSpec
		$spec.details = new-object vmware.vim.FileQueryFlags
		$spec.details.fileSize = $true
		$spec.details.fileType = $true
		$spec.details.modification = $true
		$task = $datastoreBrowser.SearchDatastoreSubFolders_Task("[$name] /", $spec)

		# Wait for the search to finish, then process the HostDatastoreBrowserSearchResults.
		# This will be a lot better in 1.0 GA.
		$view = get-view $task
		while ($view.info.state -eq "running") {
			start-sleep 1
			$view = get-view $task
		}

		foreach ($result in $view.info.result) {
			$subPath = $result.FolderPath
			foreach ($file in $result.File) {
				$fullPath = $subPath + $file.Path
				$obj = new-object psobject
				$obj | add-member -type noteproperty -name "Path" -value $fullPath
				$obj | add-member -type noteproperty -name "Size" -value $file.FileSize
				$obj | add-member -type noteproperty -name "Modification" -value $file.Modification
				$obj | add-member -type noteproperty -name "Datastore" -value $datastore.Name
				$ret += $obj
			}
		}
	}

	$ret
}

Write-Host "`tAdded Get-DatastoreFiles to global functions." -Fore White


if (Test-Path function:Register-VM) { Remove-Item function:Register-VM }
Function global:Register-VM {
	Param(
		[String[]]
		$VmxFiles,

		[VMware.VimAutomation.Client20.ResourcePoolImpl]
		$ResourcePool,

		[VMware.VimAutomation.Client20.VMHostImpl]
		$VMhost,

		[VMware.VimAutomation.Client20.FolderImpl]
		$Folder
	)

	if ($Folder -eq $null) {
		write-host -foreground yellow "Folder is required."
		return $null
	}
	$FolderView = get-view $Folder.ID

	$VMHostView = $ResourcePoolView = $null
	if ($ResourcePool) {
		$ResourcePoolView = get-view $ResourcePool.ID
	}
	if ($VMHost) {
		$VMHostView = get-view $VMHost.ID
	}

	foreach ($file in $VmxFiles) {
		$file -match '/([^/]+).vmx$' > $null
		$vmName = $matches[1]
		if (!$vmName) {
			write-host -foreground yellow "Can't determine VM name for $file, skipping"
		} else {
			$FolderView.RegisterVM($file, $vmName, $false, $ResourcePoolView.MoRef, $VMHostView.MoRef)
		}
	}
}

Write-Host "`tAdded Register-VM to global functions." -Fore White
