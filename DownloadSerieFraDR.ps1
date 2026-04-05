
param(
    [Parameter(Mandatory = $true)]
    [string]$Url,

    [Parameter(Mandatory = $false)]
    [string]$Season = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$YtDlpPath      = "C:\Users\kaspe\.stacher\yt-dlp.exe"
$FfmpegPath     = "C:\Users\kaspe\.stacher\ffmpeg.exe"
$ArchiveFile    = "C:\Users\kaspe\Downloads\YouTube\archive.txt"
$OutputRoot     = "C:\Users\kaspe\Downloads\YouTube"
$NASLocation    = "W:\TV Shows\- DR"

function Get-WebContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $headers = @{
        "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36"
        "Accept"     = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    }

    try {
        $response = Invoke-WebRequest -Uri $Url -Headers $headers -Method Get -UseBasicParsing
        if ([string]::IsNullOrWhiteSpace($response.Content)) {
            throw "Tomt svar modtaget fra $Url"
        }

        return $response.Content
    }
    catch {
        throw "Kunne ikke hente siden '$Url'. Fejl: $($_.Exception.Message)"
    }
}

function Get-ShowName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html
    )

    $patterns = @(
        '<meta\s+data-react-helmet="true"\s+property="og:title"\s+content="DRTV\s*-\s*(?<Name>[^"]+)"',
        '<title>(?<Name>[^|<]+)\|',
        '"title":"(?<Name>[^"]+)"'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match(
            $Html,
            $pattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )

        if ($match.Success) {
            return [System.Net.WebUtility]::HtmlDecode($match.Groups["Name"].Value.Trim())
        }
    }

    throw "Kunne ikke finde serienavnet i HTML-filen."
}

function Load-HtmlAgilityPack {
    $packageRoot = Join-Path $PSScriptRoot 'packages\HtmlAgilityPack'
    $dllPath = Join-Path $packageRoot 'lib\netstandard2.0\HtmlAgilityPack.dll'

    if (-not (Test-Path $dllPath)) {
        New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null

        $nupkgPath = Join-Path $packageRoot 'HtmlAgilityPack.nupkg'
        Invoke-WebRequest -Uri 'https://www.nuget.org/api/v2/package/HtmlAgilityPack' -OutFile $nupkgPath

        Expand-Archive -Path $nupkgPath -DestinationPath (Join-Path $packageRoot 'pkg') -Force

        $extractedDll = Join-Path $packageRoot 'pkg\lib\netstandard2.0\HtmlAgilityPack.dll'
        if (-not (Test-Path $extractedDll)) {
            throw 'Kunne ikke finde HtmlAgilityPack.dll i NuGet-pakken.'
        }

        New-Item -ItemType Directory -Path (Split-Path $dllPath -Parent) -Force | Out-Null
        Copy-Item -Path $extractedDll -Destination $dllPath -Force
    }

    Add-Type -Path $dllPath
}

function Convert-ToTitleCase {
    param(
        [Parameter(Mandatory)]
        [string] $Text
    )

    $cleaned = $Text.Trim().ToLowerInvariant()
    $textInfo = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo
    return $textInfo.ToTitleCase($cleaned)
}

function Get-SafeFolderName {
    param(
        [Parameter(Mandatory)]
        [string] $Name
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $safeName = $Name

    foreach ($char in $invalidChars) {
        $safeName = $safeName.Replace($char, '_')
    }

    return ($safeName -replace '\s+', ' ').Trim()
}

function Get-PlexFileBaseName {
    param(
        [Parameter(Mandatory)]
        [string] $ShowName,

        [Parameter(Mandatory)]
        [int] $SeasonNumber,

        [Parameter(Mandatory)]
        [int] $EpisodeNumber,

        [Parameter(Mandatory)]
        [string] $EpisodeTitle
    )

    $normalizedShowName = Convert-ToTitleCase -Text $ShowName
    return '{0} - S{1}E{2} - {3}' -f $normalizedShowName, $SeasonNumber, $EpisodeNumber, $EpisodeTitle
}

function Get-EpisodesFromHtml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$html,

        [Parameter(Mandatory = $true)]
        [string]$ShowName
    )

    Load-HtmlAgilityPack

    $doc = [HtmlAgilityPack.HtmlDocument]::new()
    $doc.LoadHtml($Html)

    $anchorNodes = $doc.DocumentNode.SelectNodes(
        "//a[contains(@class,'d1-drtv-episode-title-and-details')]"
    )

    if ($null -eq $anchorNodes) {
        return @()
    }

    $episodes = foreach ($anchor in $anchorNodes) {
        $url = $anchor.GetAttributeValue('href', '')
        if ([string]::IsNullOrWhiteSpace($url)) {
            continue
        }

        # erstat drtv/episode med se for at få det link der virker i yt-dlp
        $url = $url.Replace('/drtv/episode/', '/drtv/se/')

        $titleNode = $anchor.SelectSingleNode(
            ".//div[contains(@class,'d1-drtv-episode-title-and-details__contextual-title')]"
        )

        $title = if ($null -ne $titleNode) {
            [System.Net.WebUtility]::HtmlDecode($titleNode.InnerText).Trim()
        }
        else {
            ''
        }

        $episodeNumber = $title.Substring(0, 1)
        $episodeTitle = $title.Substring(3, $title.Length - 3).Trim()
        $plexTitle = Get-PlexFileBaseName -ShowName $ShowName -SeasonNumber $Season -EpisodeNumber $episodeNumber -EpisodeTitle $episodeTitle

        [PSCustomObject]@{
            ShowName      = $ShowName
            SeasonNumber  = $Season
            EpisodeNumber = $episodeNumber
            EpisodeTitle  = $episodeTitle
            EpisodeUrl    = $url
            PlexTitle     = $plexTitle
        }
    }

    $json = $episodes | ConvertTo-Json    
    return $json
}



try {
    $html = Get-WebContent -Url $Url
    $showName = Get-ShowName -Html $html
    $EpisodesJson = Get-EpisodesFromHtml -Html $html -ShowName $showName
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}

$episodes = $EpisodesJson | ConvertFrom-Json

foreach ($episode in $episodes) {
    $showFolderName = Get-SafeFolderName -Name (Convert-ToTitleCase -Text $episode.ShowName)
    $seasonFolderName = "Season $($episode.SeasonNumber)"

    $showPath = Join-Path -Path $OutputRoot -ChildPath $showFolderName
    $seasonPath = Join-Path -Path $showPath -ChildPath $seasonFolderName

    $fileBaseName = $episode.PlexTitle

    $targetPathWithoutExtension = Join-Path -Path $seasonPath -ChildPath $fileBaseName

    Write-Host "Downloader: $($episode.EpisodeUrl)"
    Write-Host "Til: $targetPathWithoutExtension"

    $outputTemplate = "$targetPathWithoutExtension.%(ext)s"

    $output = & $YtDlpPath `
        -f "bestvideo+bestaudio" `
        --verbose `
        --rm-cache-dir `
        --no-warnings `
        --embed-chapters `
        --download-archive $ArchiveFile `
        --write-sub `
        --all-subs `
        --embed-subs `
        --no-check-certificate `
        --ffmpeg-location $FfmpegPath `
        --write-subs `
        --sub-langs "da_foreign" `
        --convert-subs srt `
        -o $outputTemplate `
        --merge-output-format mkv `
        $episode.EpisodeUrl 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Download fejlede for '$($episode.PlexTitle)' med exit code $LASTEXITCODE"
    }
    elseif ($output -match "has already been recorded in the archive") {
        Write-Host "Episoden '$($episode.PlexTitle)' er allerede downloadet og flyttet til NAS tidligere."
        Write-Host ""
    }
    else {
        $NasShowLocation = "$showFolderName/$seasonFolderName"
        Write-Host "Flytter episoden til NAS: $($NasShowLocation)"
        $destinationFolder = Join-Path -Path $NASLocation -ChildPath $NasShowLocation

        # Opret destination-mappen hvis den ikke findes
        if (-not (Test-Path -LiteralPath $destinationFolder)) {
            New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        }

        # Flyt alle filer fra showPath til destinationFolder 
        Get-ChildItem -Path $seasonPath -File | ForEach-Object {
            $destinationPath = Join-Path -Path $destinationFolder -ChildPath $_.Name
            Move-Item -LiteralPath $_.FullName -Destination $destinationPath -Force
        }

        Write-Host "Færdig: $($episode.PlexTitle)"
        Write-Host ""
    }
}

# Remove the now empty series folder
Remove-Item -LiteralPath $showPath -Force -Recurse