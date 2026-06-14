# Shared jpackage helpers for package_app.ps1 and package_rdp_launcher_app.ps1
# UTF-8 BOM: Windows PowerShell 5.1

function Read-MavenPomProperties {
    param([string]$PomPath)
    [xml]$xml = Get-Content -LiteralPath $PomPath -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('m', 'http://maven.apache.org/POM/4.0.0')
    $props = @{}
    foreach ($n in $xml.SelectNodes('/m:project/m:properties/*', $ns)) {
        $props[$n.LocalName] = $n.InnerText.Trim()
    }
    return $props
}

function Expand-PomPropertyPlaceholder {
    param(
        [string]$Raw,
        [hashtable]$Props
    )
    if ($null -eq $Raw -or [string]::IsNullOrWhiteSpace($Raw)) {
        return ''
    }
    $current = $Raw.Trim()
    for ($iter = 0; $iter -lt 4; $iter++) {
        $m = [regex]::Match($current, '^\$\{([^}]+)\}$')
        if (-not $m.Success) {
            break
        }
        $innerKey = $m.Groups[1].Value
        if (-not $Props.ContainsKey($innerKey)) {
            break
        }
        $next = [string]$Props[$innerKey]
        if ([string]::IsNullOrWhiteSpace($next)) {
            break
        }
        $current = $next.Trim()
    }
    return $current
}

function Get-MavenProjectInfo {
    param([string]$PomPath)
    [xml]$xml = Get-Content -LiteralPath $PomPath -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('m', 'http://maven.apache.org/POM/4.0.0')
    $artifact = $xml.SelectSingleNode('/m:project/m:artifactId', $ns).InnerText.Trim()
    $versionNode = $xml.SelectSingleNode('/m:project/m:version', $ns)
    if (-not $versionNode -or [string]::IsNullOrWhiteSpace($versionNode.InnerText)) {
        $parentVer = $xml.SelectSingleNode('/m:project/m:parent/m:version', $ns)
        if ($parentVer) {
            $version = $parentVer.InnerText.Trim()
        }
        else {
            throw 'Could not read version from pom.xml.'
        }
    }
    else {
        $version = $versionNode.InnerText.Trim()
    }
    if (-not $artifact -or -not $version) {
        throw 'Could not read artifactId / version from pom.xml.'
    }
    $mainJar = "$artifact-$version.jar"
    @{
        ArtifactId = $artifact
        Version    = $version
        MainJar    = $mainJar
    }
}

function Copy-JpackageInputDirectory {
    param(
        [string]$RootPath,
        [string]$MainJarName,
        [string]$DestPath
    )
    if (Test-Path -LiteralPath $DestPath) {
        Remove-Item -Recurse -Force $DestPath
    }
    New-Item -ItemType Directory -Path $DestPath | Out-Null

    $mainSrc = Join-Path (Join-Path $RootPath 'target') $MainJarName
    if (-not (Test-Path -LiteralPath $mainSrc)) {
        throw "Main JAR not found: $mainSrc"
    }
    Copy-Item -LiteralPath $mainSrc -Destination $DestPath

    $depDir = Join-Path (Join-Path $RootPath 'target') 'dependency'
    if (-not (Test-Path -LiteralPath $depDir)) {
        throw "dependency folder not found: $depDir"
    }
    Copy-Item -Path (Join-Path $depDir '*') -Destination $DestPath -Force
}

function New-JavaFxJpackageModulePathJavaOptions {
    param(
        [string]$JavafxVersion,
        [string]$ModuleDirRelative = '$APPDIR\app\jfx-mod'
    )
    if ([string]::IsNullOrWhiteSpace(($JavafxVersion -replace '[\r\n\t]', '').Trim())) {
        throw 'JavafxVersion is required for jpackage module-path options.'
    }
    return @('--module-path=' + $ModuleDirRelative)
}

function Remove-DirectoryWithRetry {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxRetries = 8,
        [int]$SleepSeconds = 3,
        [string]$CloseHint = 'close Explorer / running exe on this folder'
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        return $true
    }
    for ($ri = 0; $ri -lt $MaxRetries; $ri++) {
        try {
            Remove-Item -Recurse -Force -LiteralPath $Path -ErrorAction Stop
            return $true
        }
        catch {
            Write-Warning "Cannot remove $Path ($CloseHint). Retry ($($ri + 1)/$MaxRetries)..."
            Start-Sleep -Seconds $SleepSeconds
        }
    }
    return $false
}

function Mirror-DirectoryWithRobocopy {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    & robocopy $Source $Destination /MIR /NFL /NDL /NJH /NJS /nc /ns /np /R:2 /W:2 | Out-Null
    return $LASTEXITCODE
}

function Prepare-JavaFxJpackageModuleDir {
    param(
        [Parameter(Mandatory)][string]$AppDir,
        [Parameter(Mandatory)][string]$JavafxVersion
    )
    $jv = ($JavafxVersion -replace '[\r\n\t]', '').Trim()
    $jfxDir = Join-Path $AppDir 'jfx-mod'
    if (Test-Path -LiteralPath $jfxDir) {
        Remove-Item -Recurse -Force -LiteralPath $jfxDir
    }
    New-Item -ItemType Directory -Path $jfxDir -Force | Out-Null
    $mods = @(
        'javafx-base', 'javafx-controls', 'javafx-fxml', 'javafx-graphics',
        'javafx-media', 'javafx-swing', 'javafx-web', 'jdk-jsobject')
    foreach ($prefix in $mods) {
        $leaf = "$prefix-$jv-win.jar"
        $src = Join-Path $AppDir $leaf
        if (-not (Test-Path -LiteralPath $src)) {
            throw "JavaFX module jar missing for jpackage: $src"
        }
        Copy-Item -LiteralPath $src -Destination (Join-Path $jfxDir $leaf) -Force
    }
    return $jfxDir
}

function Remove-JpackageTestArtifactsFromAppDir {
    param([Parameter(Mandatory)][string]$AppDir)
    if (-not (Test-Path -LiteralPath $AppDir)) {
        return
    }
    foreach ($pattern in @('junit*.jar', 'opentest4j*.jar', 'apiguardian*.jar')) {
        Get-ChildItem -LiteralPath $AppDir -Filter $pattern -File -ErrorAction SilentlyContinue |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    }
}

function Set-JpackageCfgJavaFxModulePathDir {
    param(
        [Parameter(Mandatory)][string]$CfgFilePath,
        [string]$ModuleDirRelative = '$APPDIR\app\jfx-mod'
    )
    if (-not (Test-Path -LiteralPath $CfgFilePath)) {
        return $false
    }
    $targetOpt = 'java-options=--module-path=' + $ModuleDirRelative
    $lines = @(Get-Content -LiteralPath $CfgFilePath -Encoding UTF8)
    if ($lines.Count -eq 0) {
        return $false
    }
    $out = [System.Collections.Generic.List[string]]::new()
    $changed = $false
    $inserted = $false
    foreach ($line in $lines) {
        if ($line -match '^java-options=--module-path=') {
            if (-not $inserted) {
                [void]$out.Add($targetOpt)
                $inserted = $true
            }
            if ($line -ne $targetOpt) {
                $changed = $true
            }
            continue
        }
        [void]$out.Add($line)
    }
    if (-not $changed) {
        return $false
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines($CfgFilePath, $out.ToArray(), $utf8NoBom)
    return $true
}

<#
  jpackage native exe reads app\*.cfg. Per-jar app.classpath + modular JavaFX breaks; match launch-pm-ai-rdp-launcher.bat:
  -classpath $APPDIR\app\*  and  --module-path=$APPDIR\app\jfx-mod
#>
function Rewrite-RdpLauncherJpackageCfgForExe {
    param(
        [Parameter(Mandatory)][string]$CfgFilePath,
        [string]$DefaultMainClass = 'jp.co.pm.ai.desktop.RemoteDesktopFxApp'
    )
    if (-not (Test-Path -LiteralPath $CfgFilePath)) {
        return $false
    }
    $existing = @(Get-Content -LiteralPath $CfgFilePath -Encoding UTF8)
    $mainClass = $DefaultMainClass
    $appVer = ''
    $xms = '512m'
    $xmx = '2g'
    $prism = 'sw'
    foreach ($line in $existing) {
        if ($line -match '^app\.mainclass=(.+)$') {
            $mainClass = $Matches[1].Trim()
        }
        elseif ($line -match '^java-options=-Djpackage\.app-version=(.+)$') {
            $appVer = $Matches[1].Trim()
        }
        elseif ($line -match '^java-options=-Xms(.+)$') {
            $xms = $Matches[1].Trim()
        }
        elseif ($line -match '^java-options=-Xmx(.+)$') {
            $xmx = $Matches[1].Trim()
        }
        elseif ($line -match '^java-options=-Dprism\.order=(.+)$') {
            $prism = $Matches[1].Trim()
        }
    }
    $out = [System.Collections.Generic.List[string]]::new()
    [void]$out.Add('[Application]')
    [void]$out.Add('app.mainclass=' + $mainClass)
    [void]$out.Add('')
    [void]$out.Add('[JavaOptions]')
    if (-not [string]::IsNullOrWhiteSpace($appVer)) {
        [void]$out.Add('java-options=-Djpackage.app-version=' + $appVer)
    }
    [void]$out.Add('java-options=-Dfile.encoding=UTF-8')
    [void]$out.Add('java-options=-Xms' + $xms)
    [void]$out.Add('java-options=-Xmx' + $xmx)
    [void]$out.Add('java-options=-Dprism.order=' + $prism)
    [void]$out.Add('java-options=-classpath')
    [void]$out.Add('java-options=$APPDIR\app\*')
    [void]$out.Add('java-options=--module-path=$APPDIR\app\jfx-mod')
    [void]$out.Add('java-options=--add-modules=javafx.controls,javafx.fxml,javafx.graphics,javafx.base,javafx.media,javafx.swing')
    [void]$out.Add('java-options=--add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED')
    [void]$out.Add('java-options=--add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED')
    [void]$out.Add('java-options=--add-opens=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED')
    [void]$out.Add('java-options=--add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED')
    [void]$out.Add('java-options=--enable-native-access=javafx.graphics')
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines($CfgFilePath, $out.ToArray(), $utf8NoBom)
    return $true
}

<#
  jpackage writes app\*.cfg with semicolon-separated java-options. A single
  --module-path=jar1;jar2 breaks into invalid lines and exe shows "Failed to launch JVM".
  Prefer Prepare-JavaFxJpackageModuleDir + Rewrite-RdpLauncherJpackageCfgForExe.
#>
function Repair-JpackageAppCfgModulePath {
    param([Parameter(Mandatory)][string]$CfgFilePath)
    if (-not (Test-Path -LiteralPath $CfgFilePath)) {
        return $false
    }
    $lines = @(Get-Content -LiteralPath $CfgFilePath -Encoding UTF8)
    if ($lines.Count -eq 0) {
        return $false
    }
    $out = [System.Collections.Generic.List[string]]::new()
    $changed = $false
    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        if ($line -eq 'java-options=--module-path=') {
            $i++
            $pathParts = [System.Collections.Generic.List[string]]::new()
            while ($i -lt $lines.Count) {
                $next = $lines[$i]
                if ($next -match '^java-options=(\$APPDIR[/\\]app[/\\].+-win\.jar(?:;.+)?)$') {
                    foreach ($part in ($Matches[1] -split ';')) {
                        $p = $part.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($p)) {
                            [void]$pathParts.Add($p)
                        }
                    }
                    $i++
                    continue
                }
                if ($next -match '^java-options=--module-path=(.+)$') {
                    foreach ($part in ($Matches[1] -split ';')) {
                        $p = $part.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($p)) {
                            [void]$pathParts.Add($p)
                        }
                    }
                    $i++
                    continue
                }
                break
            }
            if ($pathParts.Count -gt 0) {
                [void]$out.Add('java-options=--module-path=' + ($pathParts -join ';'))
                $changed = $true
            }
            else {
                [void]$out.Add('java-options=--module-path=')
                $i++
            }
            continue
        }
        [void]$out.Add($line)
        $i++
    }
    if (-not $changed) {
        return $false
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines($CfgFilePath, $out.ToArray(), $utf8NoBom)
    return $true
}

function Ensure-JdkWindowsEmbedCache {
    param(
        [string]$CacheRoot,
        [string]$JdkRelease,
        [string]$ZipUrlOverride,
        [bool]$Skip
    )

    $dest = Join-Path $CacheRoot ('jdk-embed-' + $JdkRelease + '-windows-amd64')
    $javaExe = Join-Path $dest 'bin\java.exe'
    $jpkgExe = Join-Path $dest 'bin\jpackage.exe'

    if ($Skip -and (Test-Path -LiteralPath $javaExe) -and (Test-Path -LiteralPath $jpkgExe)) {
        Write-Host "SkipJdkPrepare: using cache: $dest" -ForegroundColor DarkGray
        return [string]$dest
    }

    if (Test-Path -LiteralPath $dest) {
        Remove-Item -Recurse -Force -LiteralPath $dest
    }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    $zipPath = Join-Path $dest 'jdk-bundle.zip'
    if (-not [string]::IsNullOrWhiteSpace($ZipUrlOverride)) {
        $url = $ZipUrlOverride.Trim()
        Write-Host "--- Download JDK zip (pom pm.ai.bundle.jdk.windows.zip.url): $url ---" -ForegroundColor Cyan
    }
    else {
        $url = "https://api.adoptium.net/v3/binary/latest/$JdkRelease/ga/windows/x64/jdk/hotspot/normal/eclipse"
        Write-Host "--- Download JDK zip (Adoptium API, Windows x64 release $JdkRelease): $url ---" -ForegroundColor Cyan
    }

    Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing

    $extractTmp = Join-Path $dest '_ext'
    New-Item -ItemType Directory -Path $extractTmp -Force | Out-Null
    try {
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractTmp -Force
    }
    finally {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    }

    $javaFound = Get-ChildItem -Path $extractTmp -Recurse -Filter 'java.exe' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Directory.Name -ieq 'bin' } |
        Select-Object -First 1
    if (-not $javaFound) {
        throw "JDK zip did not contain bin\java.exe under: $extractTmp"
    }

    $jdkHome = $javaFound.Directory.Parent.FullName
    Get-ChildItem -LiteralPath $jdkHome -ErrorAction SilentlyContinue | ForEach-Object {
        Move-Item -LiteralPath $_.FullName -Destination $dest -Force
    }
    Remove-Item -LiteralPath $extractTmp -Recurse -Force -ErrorAction SilentlyContinue

    if (-not (Test-Path -LiteralPath $javaExe)) {
        throw "JDK layout error: missing $javaExe"
    }
    if (-not (Test-Path -LiteralPath $jpkgExe)) {
        throw "JDK layout error: missing $jpkgExe"
    }

    return [string]$dest
}

function Normalize-JvmHeapToken {
    param([string]$Raw)
    $t = ($Raw -replace '[\r\n\t]', '').Trim()
    if ([string]::IsNullOrWhiteSpace($t)) {
        return '512m'
    }
    return $t
}

function Sync-JavaFxWindowsRuntimeFromMavenCentral {
    param(
        [string]$PackageInputDir,
        [string]$JavafxVersion,
        [string]$CacheRoot,
        [bool]$Skip
    )

    $javaFxWinJarNames = @(
        'javafx-base',
        'javafx-controls',
        'javafx-fxml',
        'javafx-graphics',
        'javafx-media',
        'javafx-swing',
        'javafx-web',
        'jdk-jsobject'
    )
    $cacheDir = Join-Path $CacheRoot "javafx-openjfx-$JavafxVersion-windows-amd64"
    New-Item -ItemType Directory -Path $PackageInputDir -Force | Out-Null
    New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null

    foreach ($aid in $javaFxWinJarNames) {
        $fn = "$aid-$JavafxVersion-win.jar"
        $cached = Join-Path $cacheDir $fn
        $url = "https://repo1.maven.org/maven2/org/openjfx/$aid/$JavafxVersion/$fn"

        $needDownload = $true
        if ($Skip -and (Test-Path -LiteralPath $cached)) {
            $fi = Get-Item -LiteralPath $cached -ErrorAction SilentlyContinue
            if ($null -ne $fi -and $fi.Length -gt 512) {
                $needDownload = $false
                Write-Host "SkipJavaFxPrepare: using cache $fn" -ForegroundColor DarkGray
            }
        }

        if ($needDownload) {
            Write-Host "--- Download JavaFX runtime: $fn ---" -ForegroundColor Cyan
            try {
                Invoke-WebRequest -Uri $url -OutFile $cached -UseBasicParsing
            }
            catch {
                throw "JavaFX download failed: $url $($_.Exception.Message)"
            }
            $fi2 = Get-Item -LiteralPath $cached -ErrorAction SilentlyContinue
            if ($null -eq $fi2 -or $fi2.Length -lt 512) {
                throw "JavaFX download invalid or empty: $url"
            }
        }

        Copy-Item -LiteralPath $cached -Destination (Join-Path $PackageInputDir $fn) -Force
    }
}
