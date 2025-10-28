<#
.SYNOPSIS
    Script gets all User Dirs from C:\Users or C:\Benutzer etc. based on language given.
    It loops through all users and gets all Edge Extensions from File Path and exports them to a CSV:

.DESCRIPTION
    User which runs the script needs access to all %localappdata% dirs of the users on this PC. Plus he needs rwx to the csv dir and file.

.PARAMS
    $CSVPath can be added to pass the location of the destination csv file where findings should be written.
    Usage: -CSVPath "C:\Path\to\file\export.csv"


.NOTES
    Author: Jan Wiesmann
    Date: 27.10.2025
    Version: 1.0
    Script Purpose: To get a global list of edge extensions in our ecosystem. Alternative to buying MS P2 License
    Dependencies: A possibility to run script scheduled on ALL PC/Laptops. Possible from script server via WinRE or scheduled task.
#>

# Begin Script
#CSV Export Path
param(
    # Path of CSV
    [string]$CSVPath = 'C:\tmp\extensions.csv'
    ) 



#AppData Path
$appdata = $env:LOCALAPPDATA


# Resolve system user profile root dynamically (to ensure compability with multilanguage systems)
$users_path = [System.Environment]::GetFolderPath('UserProfile')
#Get User Folderes
$users_path = Split-Path $users_path -Parent
#Only take numeric ones
$Users = Get-ChildItem -Path $users_path -Name | Where-Object { $_ -match '^\d+$' }


#prepare csv output object
$Output = @()


foreach($User in $Users){
    #loop through all users in $users

    #base path of extensions
    $EdgePath = "$users_path\$User\AppData\Local\Microsoft\Edge\User Data\Default\Extensions"
    $ChromePath = "$users_path\$User\AppData\Local\Google\Chrome\User Data\Default\Extensions"

    #get all manifest.json in base path
    $paths = Get-ChildItem -Path $EdgePath -Recurse -Include *manifest.json | Select-Object FullName
    $paths += Get-ChildItem -Path $ChromePath -Recurse -Include *manifest.json | Select-Object FullName

    #loop through every manifest.json
    foreach($path in $paths){
        #get path in plaintext instead of @{FullName=xyz}
        $path = $path.FullName
        #parse manifest.json
        $Json = Get-Content $path -Raw | ConvertFrom-Json

        #Browser Name
        $browser = "Unknown"
        if ($path -match '\\Microsoft\\Edge\\') {
            $browser = "Edge"
        }
        elseif ($path -match '\\Google\\Chrome\\') {
            $browser = "Chrome"
        }
        

        #check if title of extension contains anything 
        if($Json.action.default_title -ne $null){
            $currentUser  = $User.Trim()
            $currentTitle = $Json.action.default_title.ToString().Trim()
            $currentHostname = $env:COMPUTERNAME
            # default
            $exists = $false

            #verify that path of user isn't in csv yet

            try {
                $matches = @()

                if (Test-Path $CSVPath) {
                    # look in existing CSV
                    $matches += Import-Csv -Path $CSVPath | Where-Object {
                        $_.Username.Trim() -ieq $currentUser -and
                        $_.Title.Trim()    -ieq $currentTitle -and
                        $_.Hostname.Trim() -ieq $currentHostname -and
                        $_.Browser.Trim() -ieq $browser
                    }
                }

                # also look in the in-memory $Output we've built so far
                $matches += $Output | Where-Object {
                    $_.Username.Trim() -ieq $currentUser -and
                    $_.Title.Trim()    -ieq $currentTitle -and
                    $_.Hostname.Trim() -ieq $currentHostname -and
                    $_.Browser.Trim() -ieq $browser
                }

                # exists if found in either place
                $exists = ($matches.Count -gt 0)
            }
            catch {
                $exists = $false
            }
            #If already exists it skips it to ensure no duplicates
            if (-not $exists) {
                $Output += [pscustomobject]@{
                    Hostname      = $currentHostname
                    Username      = $currentUser
                    Browser       = $browser
                    ExtensionPath = $path
                    Title         = $currentTitle
                } | Select-Object Hostname, Username, Browser, Title, ExtensionPath
            }
        }
    }
}
#export csv
$Output | Export-CSV $CSVPath -Append -NoTypeInformation 