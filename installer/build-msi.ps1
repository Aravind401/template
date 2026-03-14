param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$publishDir = Join-Path $root "publish/$Runtime"
$msiOutDir = Join-Path $root "artifacts"

Write-Host "Publishing WinForms app..."
dotnet publish "$root/QuotationTemplateApp/QuotationTemplateApp.csproj" -c $Configuration -r $Runtime --self-contained true -p:PublishSingleFile=true -o $publishDir

if (-not (Get-Command wix -ErrorAction SilentlyContinue)) {
    Write-Host "Installing WiX CLI tool..."
    dotnet tool install --global wix --version 5.*
}

if (-not (Test-Path $msiOutDir)) {
    New-Item -ItemType Directory -Path $msiOutDir | Out-Null
}

$appFiles = Get-ChildItem -Path $publishDir -File | ForEach-Object { $_.FullName } | Sort-Object
if ($appFiles.Count -eq 0) {
    throw "No published files were found in $publishDir"
}

$appFilesValue = [string]::Join(';', $appFiles)
$productWxs = Join-Path $PSScriptRoot "Product.wxs"
$outFile = Join-Path $msiOutDir "QuotationTemplateApp-$Version-$Runtime.msi"

Write-Host "Building MSI..."
wix build $productWxs -arch x64 -dAppFiles="$appFilesValue" -dVersion=$Version -o $outFile

Write-Host "MSI created: $outFile"
