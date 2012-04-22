* Nullable [DateTime]s in PowerShell:
  * Can't set null values for the DateTime properties of Package.cs
  * I *can* set Revision = $Null, though

* Out-Package won't create empty directories.
* Out-Package doesn't descend into directories passed to it.
  * Workaround for e.g. Get-Item:  use -Recurse
  * FIX THE DOCUMENTATION.

* Check ALL documentation.
* Review each cmdlet's "get an absolute path" bits and standardize if
  necessary.
