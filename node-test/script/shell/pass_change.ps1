$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)

$a = new-object -comobject wscript.shell 
$intAnswer = $a.popup("Meg szeretné változtatni az tavmeres@ipsol.hu betáplált jelszavát?", ` 
0,"Delete Files",4) 
If ($intAnswer -eq 6) {
    $actDir = $MyDir + "\creds\tavmeres-cred.txt"
    Read-Host -Prompt "Add meg az tavmeres@ipsol.hu jelszavát!" -AsSecureString | ConvertFrom-SecureString | Out-File $actDir
}