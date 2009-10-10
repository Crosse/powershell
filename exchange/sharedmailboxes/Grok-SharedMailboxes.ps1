param($infile, [switch]$stats, $permissions)

if (!($infile) -and !($permissions)) {
    Write-Error "Please specify either -infile <file> or -permissions <permissions>"
    return
}

$final = New-Object System.Collections.ArrayList

if ($infile) {
    $anyones = 0
    $others = 0
    $totalDept = 0
    $totalRequest = 0
    foreach ($line in (Get-Content $infile)) {
        if ($line.Contains("ACL:") -and ($line.Contains("dept") -or $line.Contains("request"))) {
            if ($line.Contains("dept.")) {
                $type = "dept"
            } elseif ($line.Contains("request.")) {
                $type= "request"
            }
            $temp = $line.Substring($line.IndexOf($type)).Split(" ")
            $mbox = $temp[0]
            for ($i = 1; $i -lt $temp.Count; $i+=2) { 
                if ($temp[$i].ToLower() -eq 'anyone') {
                    $anyones++
                } else {
                    $perms = $temp[$i+1].ToLower()
                    $mboxAcl = New-Object PSObject
                    $mboxAcl = Add-Member -PassThru -InputObject $mboxAcl NoteProperty SharedMailbox $mbox
                    Add-Member -InputObject $mboxAcl NoteProperty User            $temp[$i]
                    Add-Member -InputObject $mboxAcl NoteProperty Permissions     $perms
                    Add-Member -InputObject $mboxAcl NoteProperty Lookup          $perms.Contains('l')
                    Add-Member -InputObject $mboxAcl NoteProperty Read            $perms.Contains('r')
                    Add-Member -InputObject $mboxAcl NoteProperty Seen            $perms.Contains('s')
                    Add-Member -InputObject $mboxAcl NoteProperty Write           $perms.Contains('w')
                    Add-Member -InputObject $mboxAcl NoteProperty Insert          $perms.Contains('i')
                    Add-Member -InputObject $mboxAcl NoteProperty Post            $perms.Contains('p')
                    Add-Member -InputObject $mboxAcl NoteProperty Create          $perms.Contains('c')
                    Add-Member -InputObject $mboxAcl NoteProperty Delete          $perms.Contains('d')
                    Add-Member -InputObject $mboxAcl NoteProperty Administer      $perms.Contains('a')
                    $null = $final.Add($mboxAcl)
                    $others++
                }
            }
        }
    }
Write-Host "$anyones ACLs for 'anyone'`n$others ACLs for real users"
}

if ($stats) {
    if (!($permissions)) {
        $permissions = $final
    }

    $toplevels = New-Object System.Collections.Hashtable
    $subfolders = New-Object System.Collections.Hashtable
    $hwm = 0
    $hwmMailbox = ""
    foreach ($acl in $permissions) {
        if ($acl.SharedMailbox.IndexOf(".") -eq $acl.SharedMailbox.LastIndexOf(".")) {
            if ($toplevels.ContainsKey($acl.SharedMailbox)) {
                $null = $toplevels[$acl.SharedMailbox].Add($acl.User)
                if ( $toplevels[$acl.SharedMailbox].Count -gt $hwm ) {
                    $hwm = $toplevels[$acl.SharedMailbox].Count
                    $hwmMailbox = $acl.SharedMailbox
                }
            } else {
                $users = New-Object System.Collections.ArrayList
                $null = $users.Add($acl.User)
                $null = $toplevels.Add($acl.SharedMailbox, $users)
            }
        } else {
            if ($subfolders.ContainsKey($acl.SharedMailbox)) {
                $null = $subfolders[$acl.SharedMailbox].Add($acl.User)
            } else {
                $users = New-Object System.Collections.ArrayList
                $null = $users.Add($acl.User)
                $null = $subfolders.Add($acl.SharedMailbox, $users)
                if ( $subfolders[$acl.SharedMailbox].Count -gt $hwm ) {
                    $hwm = $subfolders[$acl.SharedMailbox].Count
                    $hwmMailbox = $acl.SharedMailbox
                }
            }
        }
    }
    Write-Host "$($toplevels.Count) top-level mailboxes`n$($subfolders.Count) sub-mailboxes"
    Write-Host "$($($toplevels.Count) + $($subfolders.Count)) total mailboxes"

    $mismatches = 0
    $final = New-Object System.Collections.ArrayList

    foreach ($perm in $permissions) {
        $mailbox = $perm.SharedMailbox
        if ($mailbox.Substring(5).Contains(".")) {
            $foundit = $false
            foreach ($key in $toplevels.Keys) {
                if ($key -eq $mailbox.Substring(0, $mailbox.LastIndexOf("."))) {
                    if ($toplevels[$key].Contains($perm.User)) {
                        $foundit = $true
                        break
                    }
                }
            }
            if (!($foundit)) {
                $m = New-Object PSObject
                $m = Add-Member -PassThru -InputObject $m NoteProperty Mailbox $mailbox
                Add-Member -InputObject $m NoteProperty User $perm.User
                $null = $final.Add($m)
                $mismatches++
            }
        }
    }
    Write-Host "$mismatches mismatched ACLs"
}

Write-Output $final
