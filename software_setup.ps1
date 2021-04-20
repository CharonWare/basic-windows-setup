$hostname = $env:COMPUTERNAME

Write-Host "$hostname"

# Import the ActiveDirectory module because a standard windows image will not have this, it is required to later add the PC to the security group

# Server needs to have ADDS and ADWS running to import modules from it
$server = New-PSSession -Computer $ADserver -Credential domain\admin
Invoke-Command -Session $server -ScriptBlock {Import-Module ActiveDirectory}
Import-PSSession -Session $server -Module ActiveDirectory -Prefix REM

$working_dir = "\\server\software"
Write-Host "Moving to $working_dir"

Set-Location -Path $working_dir

# This function is used to cut down on the re-use of this code
function msi_install($software_string){

    Start-Process msiexec.exe -Wait -ArgumentList $software_string

}

# No error handling here, but powershell will print an error if something goes wrong

msi_install('/I GoogleChromeStandaloneEnterprise64.msi /quiet')
Write-Host "Chrome installed."

msi_install('/i wsasme.msi GUILIC=<webroot keycode> CMDLINE=SME,quiet /qn /l*v install.log')
Write-Host "Webroot installed."


# Function created for openoffice since it requires files beyond the msi

function OpenOffice{
    
    $response = Read-Host -Prompt "Would you like to install OpenOffice? y/n"
    if ($response -eq "y")
    {
        Set-Location -Path "$working_dir\Openoffice"
        msi_install('/i "openoffice419.msi" /q /norestart SETUP_USED=1 RebootYesNo=No CREATEDESKTOPLINK=1 ADDLOCAL=ALL')
        Set-Location $working_dir
        Write-Host "OpenOffice installed."
    }
    Else 
    {
        Continue
    }
}

OpenOffice

# Add the computer to the security group to ensure that it picks up all the GPOs
ADD-ADGroupMember "Domain_PCs" -members $Hostname

Write-Host "Restarting PC"
Restart-Computer -Confirm