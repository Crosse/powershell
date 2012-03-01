$assembly = Get-ChildItem -Path 'C:\Program Files\Reference Assemblies\Microsoft\Framework' -Filter "WindowsBase.dll" -Recurse
if ($assembly -eq $null) {
    throw New-Object System.IO.FileNotFoundException "Cannot find WindowsBase.dll"
}
Add-Type -Path $assembly.FullName
