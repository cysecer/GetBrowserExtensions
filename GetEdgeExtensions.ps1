#AppData Path
$appdata = $env:LOCALAPPDATA
#prepare csv output object
$Output = @()
#CSV Path
$csv_path = "<path_to_csv>\extensions.csv" #might use network drive if script is ran on all pc

#base path of extensions
$BasePath = "$appdata\Microsoft\Edge\User Data\Default\Extensions"

#get all manifest.json in base path
$paths = Get-ChildItem -Path $BasePath -Recurse -Include *manifest.json | Select-Object FullName

#loop through every manifest.json
foreach($path in $paths){
    #get path in plaintext instead of @{FullName=xyz}
    $path = $path.FullName
    #parse manifest.json
    $Json = Get-Content $path -Raw | ConvertFrom-Json


    

    #check if title of extension contains anything 
    if($Json.action.default_title -ne $null){

        #verify that path of user isn't in csv yet

        if(Test-Path $csv_path){
            #CSV Exists
            try{ #try to read it and find out if entry exists already
                 #
                 #Save Array from CSV and search for Username + Title of Extension
                 $matches = [array](Import-Csv -Path $csv_path | Where-Object {
                    $_.Username.Trim() -ieq $User.ToString().Trim() -and
                    $_.Title.Trim() -ieq $Json.action.default_title.ToString().Trim()
                 })
                 #If Exists $exists gets true
                 $exists = ($matches.Count -gt 0)
            }
            catch{ # if couldn't be read
                Write-Host "File exists but could not be read: $($_.Exception.Message)"
            }
        }else{
             #CSV Not found - probably wasn't created yet.
             $exists = $false #to ensure it will be added to csv
        }

        #If already exists it skips it to ensure no duplicates
        if($exists -ne $true){
            #add to csv
            $Output += New-Object -TypeName PSObject -Property @{
                Username = $User
                ExtensionPath = $path
                Title = $Json.action.default_title
            } | Select-Object Username, Title, ExtensionPath
        }
    }
}

#export csv
$Output | Export-CSV $csv_path -Append -NoTypeInformation 
