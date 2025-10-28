<#
.SYNOPSIS
    Script gets all User Dirs from C:\Users or C:\Benutzer etc. based on language given.
    It loops through all users and gets all Edge, Chrome and Firefox Extensions from File Path and exports them to a CSV:

.DESCRIPTION
    User which runs the script needs access to all %localappdata% and %appdata% dirs of the users on this PC. Plus he needs rwx to the csv dir and file.

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






# Function: Get-FirefoxUserExtensions -User "Username"
# Goal: Get all Extensions from Firefox
# Will be called during loop of main script. Is solved with a separate function because of space reasons.
function Get-FirefoxUserExtensions {
  [CmdletBinding()]
  param(
    #Require Username Param
    [Parameter(Position=0,mandatory=$true)]
    [string]$Username,
    # Optional switch: include disabled extensions too
    [Parameter(Position=1)]
    [switch]$IncludeDisabled
  )

  #Get Users Folder (To ensure German / English Path support - C:\Users or C:\Benutzer)
  $userpath = [System.Environment]::GetFolderPath('UserProfile')
  #Remove Username from Path to get just Users Folder
  $userpath = Split-Path $userpath -Parent

  # Get the path to Firefox's profiles.ini for User $Username (contains all user profile paths)
  $profilesIni = Join-Path $userpath "\$Username\AppData\Roaming\Mozilla\Firefox\profiles.ini"

  # Verify the file exists (Firefox must be installed)
  if (-not (Test-Path $profilesIni)) {
    Write-Error ("profiles.ini not found at {0}. Is Firefox installed for this user?" -f $profilesIni)
    return
  }

  # Read the entire INI file
  $iniText = Get-Content $profilesIni -Raw
  # Split into sections [Profile0], [Profile1], etc.
  $sections = $iniText -split '(?=\[Profile\d+\])'

  # Helper function to safely retrieve the add-on name
  function Get-AddonName([object]$a) {
    if ($a -and $a.defaultLocale -and $a.defaultLocale.name) { return $a.defaultLocale.name }
    if ($a -and $a.name) { return $a.name }
    return $null
  }

  # Helper function: detect Mozilla’s built-in extensions by their ID
  $isMozillaId = {
    param($id)
    if (-not $id) { return $false }
    # Match any extension IDs ending with @mozilla.org, @mozilla.com, or @search.mozilla.org
    return ($id -match '@mozilla\.(org|com)$' -or $id -match '@search\.mozilla\.org$')
  }

  # Collect results for all profiles
  $results = foreach ($s in $sections) {
    # Skip any section that doesn’t define a profile
    if ($s -notmatch '\[Profile\d+\]') { continue }
    # Extract the Path line from the section
    if ($s -notmatch 'Path=(.+)') { continue }

    $path  = $Matches[1].Trim()
    # Determine whether the path is relative or absolute
    $isRel = if ($s -match 'IsRelative=(\d)') { $Matches[1] -eq '1' } else { $true }
    # Build the absolute path to the profile folder
    $profileDir = if ($isRel) { Join-Path (Split-Path $profilesIni) $path } else { $path }

    # Locate the extensions.json file (stores all extension info for this profile)
    $extJson = Join-Path $profileDir 'extensions.json'
    if (-not (Test-Path $extJson)) { continue }

    # Try to parse the JSON file
    try {
      $json = Get-Content $extJson -Raw | ConvertFrom-Json
    } catch {
      # Use formatted string to avoid colon variable parsing errors
      Write-Warning ("Failed to parse {0}: {1}" -f $extJson, $_)
      continue
    }

    # Filter only the user-installed extensions:
    # - Must be of type "extension"
    # - Must be installed in the user's profile directory ("app-profile")
    # - Must NOT be a Mozilla or search engine built-in
    $addons = $json.addons | Where-Object {
      $_.type -eq 'extension' -and
      $_.location -eq 'app-profile' -and
      -not (& $isMozillaId $_.id)
    }

    # If the user didn’t request disabled ones, filter out inactive extensions
    if (-not $IncludeDisabled) { $addons = $addons | Where-Object { $_.active } }

    # Save all Infos in a Object
    foreach ($a in $addons) {
      [pscustomobject]@{
        Hostname         = $env:COMPUTERNAME
        Username         = $Username
        Browser          = "Mozilla Firefox"
        Title            = Get-AddonName $a     
        ExtensionPath    = $a.rootURI    
      }
    }
  }

  # Return object
  return $results
}

################################ MAIN ################################

#AppData Path
$appdata = $env:LOCALAPPDATA


# Resolve system user profile root dynamically (to ensure compability with multilanguage systems)
$users_path = [System.Environment]::GetFolderPath('UserProfile')
#Get User Folderes
$users_path = Split-Path $users_path -Parent
#Only take numeric ones
$Users = Get-ChildItem -Path $users_path -Name # If only numeric usernames are relevant add: | Where-Object { $_ -match '^\d+$' }


#prepare csv output object
$Output = @()


foreach($User in $Users){
    #loop through all users in $users

    #base path of extensions
    $EdgePath = "$users_path\$User\AppData\Local\Microsoft\Edge\User Data\Default\Extensions"
    $ChromePath = "$users_path\$User\AppData\Local\Google\Chrome\User Data\Default\Extensions"

    ##########################################################################################################################################
    #Get Firefox Extensions
    $Firefox = @()
    $Firefox = Get-FirefoxUserExtensions -Username $User
    
    foreach ($row in $Firefox) {
        $currentUser      = $row.Username.Trim()
        $currentTitle     = $row.Title.ToString().Trim()
        $currentHostname  = $row.Hostname.Trim()
        $browser          = $row.Browser  

        $exists = $false

        try {
            $matches = @()

            # look in existing CSV
            if (Test-Path $CSVPath) {
                $matches += Import-Csv -Path $CSVPath | Where-Object {
                    $_.Username.Trim() -ieq $currentUser -and
                    $_.Title.Trim()    -ieq $currentTitle -and
                    $_.Hostname.Trim() -ieq $currentHostname -and
                    $_.Browser.Trim()  -ieq $browser
                }
            }

            # also look in in-memory $Output built so far
            $matches += $Output | Where-Object {
                $_.Username.Trim() -ieq $currentUser -and
                $_.Title.Trim()    -ieq $currentTitle -and
                $_.Hostname.Trim() -ieq $currentHostname -and
                $_.Browser.Trim()  -ieq $browser
            }

            $exists = ($matches.Count -gt 0)
        }
        catch {
            $exists = $false
        }

        if (-not $exists) {
            # append to your accumulator
            $Output += $row
        }
    }

    ##########################################################################################################################################
    # Get Edge + Chrome Extensions

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

