Function ConvertFrom-SecureString-AsPlainText{
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true
        )]
        [System.Security.SecureString]
        $SecureString
    )
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString);
    $PlainTextString = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr);
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    $PlainTextString;
}

<#
$Varriable = ConvertFrom-SecureString-AsPlainText -SecureString (get-content <PathToTxtKey>\<Credential>.txt | convertto-securestring)
#>
