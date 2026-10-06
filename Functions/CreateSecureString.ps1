### Creates the secure password as a file ####

$PathToFolderWithCredentials = "<PathToStoredCrednetials>\StoredCredentials\"

#write-host "Enter login as domain\login:"
#read-host | out-file $PathToFolderWithCredentials\login.txt

write-host "Enter password:"
$SecurestringConverted = read-host -assecurestring | convertfrom-securestring
write-host "Enter file name:"
$filename = read-host
$Outpath = $PathToFolderWithCredentials+$filename+".txt"

$Securestringconverted | out-file $outpath

write-host "*** Credentials have been saved to $pathtofolderwithcredentials ***"

<#
### Using the file ####

$login= get-content $PathToFolderWithCredentials\login.txt
$password = get-content $PathToFolderWithCredentials\pass.txt | convertto-securestring
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $login,$password

Note: Passwords saved with ConvertFrom-SecureString (without a custom key) are encrypted using Windows Data Protection API (DPAPI).
The resulting file can only be decrypted by the same user account on the same computer that created it.
#>
