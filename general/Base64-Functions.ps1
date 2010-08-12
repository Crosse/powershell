if (Test-Path function:ConvertTo-Base64) { 
    Remove-Item function:ConvertTo-Base64
}
function global:ConvertTo-Base64($string) {
   $bytes  = [System.Text.Encoding]::UTF8.GetBytes($string);
   $encoded = [System.Convert]::ToBase64String($bytes); 

   return $encoded;
}


if (Test-Path function:ConvertFrom-Base64) { 
    Remove-Item function:ConvertFrom-Base64
}
function global:ConvertFrom-Base64($string) {
   $bytes  = [System.Convert]::FromBase64String($string);
   $decoded = [System.Text.Encoding]::UTF8.GetString($bytes); 

   return $decoded;
}

