$companyNameValue = 'Ryan LLC'

# Step 1: Read package_version_template.ini
$packageVersioningTemplate = Get-Content -Raw -Path "package_version_template.ini"

# Step 2: Replace '%%CompanyName%%' with $companyNameValue
$packageVersioningTemplate = $packageVersioningTemplate -replace '%%CompanyName%%', $companyNameValue

# Step 3: Replace '%%Year%%' with the current year
$currentYear = (Get-Date).Year
$packageVersioningTemplate = $packageVersioningTemplate -replace '%%Year%%', $currentYear

# Step 4 & 5: Handle version.ini
$versionFile = "version.ini"
$major = 1; $minor = 0; $build = 0
if (Test-Path $versionFile) {
    $versionLine = Select-String -Path $versionFile -Pattern 'filevers=\((\d+),\s*(\d+),\s*(\d+)' | Select-Object -First 1
    if ($versionLine) {
        $matches = [regex]::Match($versionLine.Line, 'filevers=\((\d+),\s*(\d+),\s*(\d+)')
        $major = [int]$matches.Groups[1].Value
        $minor = [int]$matches.Groups[2].Value
        $build = [int]$matches.Groups[3].Value
        $suggested = "$major.$minor." + ($build + 1)
        $prompt = "Last version was v$major.$minor.$build. Type in a new version number or press enter to accept v$suggested)"
        $inputVersion = Read-Host $prompt
        if ([string]::IsNullOrWhiteSpace($inputVersion)) {
            $inputVersion = $suggested
        }
    } else {
        $inputVersion = Read-Host "Type in a new version number in x.y.z format"
    }
} else {
    $inputVersion = Read-Host "No pre-existing version.ini found. Type in a new version number in x.y.z format"
}

# Step 6: Parse and replace version numbers
if ($inputVersion -match '^(\d+)\.(\d+)\.(\d+)$') {
    $major = $matches[1]
    $minor = $matches[2]
    $build = $matches[3]
} else {
    Write-Host "Invalid version format. Please use x.y.z"
    exit 1
}

$packageVersioningTemplate = $packageVersioningTemplate `
    -replace '%%Major%%', $major `
    -replace '%%Minor%%', $minor `
    -replace '%%Build%%', $build

# Step 7: Write to version.txt
Set-Content -Path $versionFile -Value $packageVersioningTemplate

Write-Host -Fore Cyan "version.txt updated to v$major.$minor.$build"