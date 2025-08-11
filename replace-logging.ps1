# PowerShell script na nahradenie logging metód v balíku
# Používa iba Abstractions-compatible helper metódy

Write-Host "Replacing logging methods in package files..."

# Získaj všetky .cs súbory v balíku (nie demo)
$files = Get-ChildItem -Path "RpaWinUiComponentsPackage\RpaWinUiComponentsPackage" -Filter "*.cs" -Recurse

foreach ($file in $files) {
    # Vynechaj LoggerExtensions súbory
    if ($file.Name -eq "LoggerExtensions.cs") {
        continue
    }
    
    $content = Get-Content $file.FullName -Raw
    $originalContent = $content
    
    # Replace all logging variants with unified helper methods
    $content = $content -replace '\.LogInformation\(', '.Info('
    $content = $content -replace '\.LogInfo\(', '.Info('
    $content = $content -replace '\.LogError\(', '.Error('
    $content = $content -replace '\.LogErr\(', '.Error('  
    $content = $content -replace '\.LogDebug\(', '.Debug('
    $content = $content -replace '\.LogDbg\(', '.Debug('
    $content = $content -replace '\.LogWarning\(', '.Warning('
    $content = $content -replace '\.LogWarn\(', '.Warning('
    
    # Ak sa obsah zmenil, zapíš naspäť
    if ($content -ne $originalContent) {
        Set-Content -Path $file.FullName -Value $content -NoNewline
        Write-Host "Updated: $($file.Name)"
    }
}

Write-Host "Logging method replacement completed!"