
$YtDlpPath      = "C:\Users\kaspe\.stacher\yt-dlp.exe"
$FfmpegPath     = "C:\Users\kaspe\.stacher\ffmpeg.exe"
$ArchiveFile    = "C:\Users\kaspe\Downloads\YouTube\archive.txt"
$OutputRoot     = "C:\Users\kaspe\Downloads\YouTube"
$NASLocation    = "W:\Movies\- DR"

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function Get-DrtvFilmsFromEmbeddedJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html
    )

    $results = New-Object System.Collections.Generic.List[object]

    # Matcher title + watchPath i de indlejrede dataobjekter
    $pattern = '"title":"(?<title>(?:\\.|[^"\\])+)".*?"watchPath":"(?<watchPath>(?:\\.|[^"\\])+)"'

    $titleMatches = [regex]::Matches(
        $Html,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    foreach ($match in $titleMatches) {
        $title = [regex]::Unescape($match.Groups['title'].Value)
        $watchPath = [regex]::Unescape($match.Groups['watchPath'].Value)

        if ([string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($watchPath)) {
            continue
        }

        # Filter our titles that are not actual movies but sections or categories on the DR website
        if ($title -eq "Anmelderroste film" -or $title -eq "DRTV - Stream TV online her" -or $title -eq "Film | Netop tilføjet") {
            Write-Host "Udeladt film: $title"
            continue
        }

        $fullLink =
            if ($watchPath -match '^https?://') {
                $watchPath
            }
            elseif ($watchPath.StartsWith('/')) {
                "https://www.dr.dk/drtv$watchPath"
            }
            else {
                "https://www.dr.dk/drtv/$watchPath"
            }
        
        Write-Host "Fundet film: $title"
        $results.Add([pscustomobject]@{
            name = $title
            link = $fullLink
        })
    }

    return $results
}

function Get-DrtvFilmsFromAnchors {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html
    )

    $results = New-Object System.Collections.Generic.List[object]

    # Fallback: finder filmkort med aria-label + href
    $pattern = '<a[^>]*class="[^"]*packshot[^"]*"[^>]*aria-label="(?<title>[^"]+)"[^>]*href="(?<href>[^"]+)"'

    $movieMatches = [regex]::Matches(
        $Html,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    foreach ($match in $movieMatches) {
        $title = $match.Groups['title'].Value
        $href  = $match.Groups['href'].Value

        if ([string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($href)) {
            continue
        }

        $fullLink =
            if ($href -match '^https?://') {
                $href
            }
            elseif ($href.StartsWith('/')) {
                "https://www.dr.dk$href"
            }
            else {
                "https://www.dr.dk/$href"
            }

        $results.Add([pscustomobject]@{
            name = $title
            link = $fullLink
        })
    }

    return $results
}

function Remove-Duplicates {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Items
    )

    return $Items |
        Group-Object -Property name, link |
        ForEach-Object { $_.Group[0] } |
        Sort-Object name
}

function Convert-ToSafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $safeName = $Name

    foreach ($char in $invalidChars) {
        $safeName = $safeName.Replace($char, "_")
    }

    $safeName = $safeName.Trim()

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw "Navnet '$Name' kunne ikke konverteres til et gyldigt filnavn."
    }

    return $safeName
}

function Assert-PathExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Description blev ikke fundet: $Path"
    }
}

function FetchFromUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $html = Get-WebContent -Url $Url

    $films = Get-DrtvFilmsFromEmbeddedJson -Html $html

    if ($films.Count -eq 0) {
        $films = Get-DrtvFilmsFromAnchors -Html $html
    }

    $uniqueFilms = Remove-Duplicates -Items $films

    if ($uniqueFilms.Count -eq 0) {
        throw "Ingen film blev fundet på siden. DR har sandsynligvis ændret markup eller loader data på en anden måde."
    }

    Assert-PathExists -Path $YtDlpPath -Description "yt-dlp"
    Assert-PathExists -Path $FfmpegPath -Description "ffmpeg"

    if (-not (Test-Path -LiteralPath $OutputRoot)) {
        New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
    }

    $archiveDirectory = Split-Path -Path $ArchiveFile -Parent
    if (-not (Test-Path -LiteralPath $archiveDirectory)) {
        New-Item -ItemType Directory -Path $archiveDirectory -Force | Out-Null
    }

    $items = $uniqueFilms

    if ($null -eq $items -or $items.Count -eq 0) {
        throw "JSON-filen indeholder ingen elementer."
    }

    foreach ($item in $items) {
        if ([string]::IsNullOrWhiteSpace($item.name)) {
            Write-Warning "Springer element over, fordi 'name' mangler."
            continue
        }

        if ([string]::IsNullOrWhiteSpace($item.link)) {
            Write-Warning "Springer '$($item.name)' over, fordi 'link' mangler."
            continue
        }

        $movieName = Convert-ToSafeFileName -Name $item.name
        $movieFolder = Join-Path -Path $OutputRoot -ChildPath $movieName
        $outputTemplate = Join-Path -Path $movieFolder -ChildPath "$movieName.%(ext)s"

        Write-Host "Downloader: $($item.name)"
        Write-Host "Link: $($item.link)"
        Write-Host "Mappe: $movieFolder"
        Write-Host ""

        $output = & $YtDlpPath `
            -f "bestvideo+bestaudio" `
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
            $item.link 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Download fejlede for '$($item.name)' med exit code $LASTEXITCODE"
        }
        elseif ($output -match "has already been recorded in the archive") {
            Write-Host "Filmen '$($item.name)' er allerede downloadet og flyttet til NAS tidligere."
            Write-Host ""
        }
        else {
            Write-Host "Flytter filmen til NAS: $($item.name)"
            $destinationFolder = Join-Path -Path $NASLocation -ChildPath $movieName

            # Opret destination-mappen hvis den ikke findes
            if (-not (Test-Path -LiteralPath $destinationFolder)) {
                New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
            }

            # Flyt alle filer fra movieFolder til destinationFolder med progressbar output
            Get-ChildItem -Path $movieFolder -File | ForEach-Object {
                $destinationPath = Join-Path -Path $destinationFolder -ChildPath $_.Name
                Move-Item -LiteralPath $_.FullName -Destination $destinationPath -Force
            }

            # Remove the now empty movie folder
            Remove-Item -LiteralPath $movieFolder -Force

            Write-Host "Færdig: $($item.name)"
            Write-Host ""
        }
    }
}

Write-Host "Henter seneste film fra DR..."
FetchFromUrl -Url "https://www.dr.dk/drtv/liste/film-_-netop-tilfoejet_409048"

Write-Host "Henter anmelderroste film fra DR..."
FetchFromUrl -Url "https://www.dr.dk/drtv/liste/anmelderroste-film_351892"

Write-Host "Henter mest sete film fra DR..."
FetchFromUrl -Url "https://www.dr.dk/drtv/liste/ta-film-_-mest-sete_520856"

Write-Host "Henter filmklassikere fra DR..."
FetchFromUrl -Url "https://www.dr.dk/drtv/liste/kategorier_film_release-year_filmklassikere_501049"

Write-Host "Henter krimi og thrillere fra DR..."
FetchFromUrl -Url "https://www.dr.dk/drtv/liste/kategorier_film_krimi_thriller_481440"