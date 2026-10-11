param([string]$FrameworkRootPath)

$ErrorActionPreference = 'Stop'
$buildTools = Join-Path $PSScriptRoot '.build-tools'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$msbuild = & $vswhere -latest -products '*' -find 'MSBuild/Current/Bin/MSBuild.exe' | Select-Object -First 1
if (!$msbuild) { throw 'Install Visual Studio with .NET desktop development first.' }

$installedReferences = Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies/Microsoft/Framework/.NETFramework/v4.6.2'
if (!$FrameworkRootPath -and !(Test-Path -LiteralPath $installedReferences)) {
    $referencePackage = Join-Path $buildTools 'net462.nupkg'
    $referenceFolder = Join-Path $buildTools 'net462'
    $FrameworkRootPath = Join-Path $referenceFolder 'build'
    if (!(Test-Path -LiteralPath $FrameworkRootPath)) {
        New-Item -ItemType Directory -Path $buildTools -Force | Out-Null
        Invoke-WebRequest -Uri 'https://api.nuget.org/v3-flatcontainer/microsoft.netframework.referenceassemblies.net462/1.0.3/microsoft.netframework.referenceassemblies.net462.1.0.3.nupkg' -OutFile $referencePackage
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($referencePackage, $referenceFolder)
    }
}

$options = @('/nologo', '/m:1', '/nr:false', '/restore', '/p:UseSharedCompilation=false', '/p:Configuration=Release', '/p:Platform=AnyCPU',
    '/p:SignAssembly=false', "/p:SolutionDir=$PSScriptRoot/", "/p:RestorePackagesPath=$buildTools/packages")
if ($FrameworkRootPath) { $options += "/p:TargetFrameworkRootPath=$FrameworkRootPath/" }

Push-Location $PSScriptRoot
try {
    foreach ($project in @('src/clawPDF/clawPDF.csproj', 'src/clawPDF.Bridge/clawPDF.Bridge.csproj', 'src/SetupHelper/SetupHelper.csproj', 'src/clawPDF.WordAddIn/clawPDF.WordAddIn.csproj')) {
        & $msbuild $project @options /v:minimal
        if ($LASTEXITCODE -ne 0) { throw "Build failed: $project" }
    }
    Write-Host "Unsigned build: $PSScriptRoot/src/_Build/Release"
} finally {
    Pop-Location
}
