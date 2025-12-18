Clear-Host

# --- CONFIG ---
$LibraryName           = 'Movies'
$TotalCount            = 10
$CachePath             = Join-Path $env:LOCALAPPDATA 'Plex\localServer.json'


Function Save-PlexLocalServerCache{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [pscustomobject]$LocalServer,

        [Parameter(Mandatory)]
        [string]$Path
    )
    Try{
        New-Item -ItemType Directory -Path (Split-Path $Path) -Force | Out-Null
    }
    Catch{
        Throw "Failed to provision a directory for locally cached server details: $($_.Exception.Message)"
    }
    Try{
        $LocalServer | ConvertTo-Json -Depth 3 | Set-Content -Path $Path -Encoding UTF8 | Out-Null
    }
    Catch{
        Throw "Failed to save locally cached server details: $($_.Exception.Message)"
    }
}

Function Load-PlexLocalServerCache{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    If(!(Test-Path $Path)) {
        Return $null
    }
    Try{
        Get-Content $Path -Raw | ConvertFrom-Json
    }
    Catch{
        return $null
    }
}

Function Get-PlexTokenFromCloud{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$AccountEmailAddress,

        [Parameter(Mandatory)]
        [string]$AccountPassword
    )

    # --- LOGIN TO PLEX CLOUD ---
    $headers = @{
        "X-Plex-Client-Identifier" = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Cryptography").MachineGuid
        "X-Plex-Product"           = "TokenFetcher"
        "X-Plex-Version"           = "1.0"
        "X-Plex-Device"            = "Windows"
        "Content-Type"             = "application/x-www-form-urlencoded"
    }
    Try{
        $loginParams = @{
            Method      = 'Post'
            Uri         = 'https://plex.tv/users/sign_in.json'
            Headers     = $headers
            Body        = "user[login]=$AccountEmailAddress&user[password]=$AccountPassword"
            ErrorAction = 'Stop'
        }
        $login = Invoke-RestMethod @loginParams
        $cloudToken = $login.user.authentication_token
    }
    Catch{
        Throw "Failed to connect to Plex Cloud: $($_.Exception.Message)"
    }
    If(!$cloudToken){
        Throw "Plex Cloud did not return an authentication token"
    }
    Return $cloudToken;
}

Function Get-PlexLocalServerFromCloud{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$CloudToken
    )

    # --- GET SERVERS YOU HAVE ACCESS TO ---
    Try{
        $resourceParams = @{
            Method      = 'Get'
            Uri         = "https://plex.tv/api/resources?includeHttps=1&includeRelay=1&X-Plex-Token=$CloudToken"
            ErrorAction = 'Stop'
        }
        $resources = Invoke-RestMethod @resourceParams
    }
    Catch{
        Throw "Failed to get local Plex servers: $($_.Exception.Message)"
    }
    $servers = $resources.MediaContainer.Device | Where-Object{ $_.provides -Match 'server' }
    If(!$servers){
        Throw "No local Plex servers available"
    }

    # --- GET SERVERS YOU OWN ---
    $ownedServer = $servers | Where-Object{ $_.owned -Eq 1 } | Select-Object -First 1
    If(!$ownedServer){
        Throw "No owned Plex servers available"
    }

    # --- GET THE CONNECTION FOR YOUR OWNED SERVER ---
    $connection = $ownedServer.Connection | Where-Object{ $_.local -Eq 1 -And -Not $_.relay -And $_.port -Eq '32400' } | Select-Object -First 1
    If(!$connection){
        Throw "No valid local connection found for the owned Plex server"
    }
    Return [PSCustomObject]@{
        AccessToken = $ownedServer.accessToken
        Name        = $ownedServer.name
        Uri         = $connection.uri
    }
}


Function Get-PlexRecentlyAddedMovieItems{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,

        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter(Mandatory)]
        [string]$LibraryName,

        [Parameter(Mandatory)]
        [int]$Count
    )

    $headers = @{ Accept = 'application/xml' }

    # --- GET LIBRARY SECTIONS ---
    Try{
        $sectionParams = @{
            Uri         = "$ServerUrl/library/sections?X-Plex-Token=$AccessToken"
            Headers     = $headers
            ErrorAction = 'Stop'
        }
        $sections = Invoke-RestMethod @sectionParams
    }
    Catch{
        Throw "Failed to get libraries from local Plex server: $($_.Exception.Message)"
    }
    $requiredSection = $sections.MediaContainer.Directory | Where-Object{ $_.type -Eq 'movie' -And $_.title -Eq $LibraryName } | Select-Object -First 1
    If(!$requiredSection){
        Throw "Failed to find a library named '$LibraryName'"
    }

    # --- GET RECENTLY ADDED MOVIES ---
    Try{
        $recentParams = @{
            Uri         = "$ServerUrl/library/sections/$($requiredSection.key)/all?type=1&sort=addedAt:desc&X-Plex-Container-Start=0&X-Plex-Container-Size=$Count&X-Plex-Token=$AccessToken"
            Headers     = $headers
            ErrorAction = 'Stop'
        }
        $recentlyAddedItems = Invoke-RestMethod @recentParams
    }
    Catch{
        Throw "Failed to query the library '$LibraryName' for recent additions: $($_.Exception.Message)"
    }

    $recentlyAddedVideos = $recentlyAddedItems.MediaContainer.Video
    If(!$recentlyAddedVideos){
        Throw "There are no recently added videos in '$LibraryName'"
    }

    # --- GET MOVIE DETAILS PER VIDEO ---
    $recentlyAddedMovies = New-Object System.Collections.Generic.List[object]($Count)
    ForEach($recentlyAddedVideo In ($recentlyAddedVideos | Select-Object -First $Count)){

        $recentlyAddedMovies.Add(
            [PSCustomObject]@{
                Title   = $recentlyAddedVideo.title
                Year    = $recentlyAddedVideo.year
                Thumb   = "$ServerUrl$($recentlyAddedVideo.thumb)?X-Plex-Token=$AccessToken"
                AddedAt = [DateTimeOffset]::FromUnixTimeSeconds([int64]$recentlyAddedVideo.addedAt).ToLocalTime()
                Slug    = $recentlyAddedVideo.slug
                Key     = $recentlyAddedVideo.ratingKey
            }
        )
    }
    Return $recentlyAddedMovies
}


# --- MAIN EXECUTION ---
$MyLocalServer = Load-PlexLocalServerCache `
    -Path $CachePath

If(!$MyLocalServer){
    Try{
        $PlexServerCreds = Get-Credential -Title "No locally cached connection found, so will need to authenticate via Plex Cloud" -Message "To do this, you need to provide email/password access - these are used one time and are not stored:"
        $MyPlexCloudToken = Get-PlexTokenFromCloud `
            -AccountEmailAddress $PlexServerCreds.UserName `
            -AccountPassword     ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PlexServerCreds.Password)))
    }
    Catch{
        Throw "Unknown error retrieving cloud token: $($_.Exception.Message)"
    }
    Try{
        $MyLocalServer = Get-PlexLocalServerFromCloud `
            -CloudToken $MyPlexCloudToken
    }
    Catch{
        Throw "Unknown error retrieving local server: $($_.Exception.Message)"
    }
    Save-PlexLocalServerCache `
        -LocalServer $MyLocalServer `
        -Path $CachePath
}

Try{
    $MyRecentMovies = Get-PlexRecentlyAddedMovieItems `
        -ServerUrl   $MyLocalServer.Uri `
        -AccessToken $MyLocalServer.AccessToken `
        -LibraryName $LibraryName `
        -Count       $TotalCount
}
Catch{
    Throw "Unknown error retrieving recent movies: $($_.Exception.Message)"
}

$MyRecentMovies | Format-Table
