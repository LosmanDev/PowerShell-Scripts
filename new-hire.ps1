# Load required assemblies
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


$basePaths = @{
    # Source documents replace username with your own.
    WelcomeLetterTemplate = "C:\"
    SecureLetterTemplate  = "C:\"
    
    DesktopRoot           = "C:\"
    DownloadsRoot         = "C:\"
    
    SecureEmailGuide      = "C:\"
    AutoPilotGuide        = "C:\"
    EmailTemplate         = "C:\"
    
    # URLs
    IntuneGroupURL        = ""
    SmartsheetURL         = ""
    
    # Filename templates
    SecurePDFNameTemplate = ""
    # FedEx pattern
    FedExPattern          = "FedEx-Shipping-Label"
}
#endregion

# Function to convert Word to PDF
function ConvertTo-PDF {
    param (
        [string]$WordPath,
        [string]$PDFPath
    )
    
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($WordPath)
    $doc.SaveAs([ref]$PDFPath, [ref]17)  # 17 = wdFormatPDF
    $doc.Close()
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

# Function to set text in Word document
function Set-WordText {
    param (
        [string]$DocPath,
        [hashtable]$Replacements
    )
    
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($DocPath)
    
    foreach ($key in $Replacements.Keys) {
        $find = $word.Selection.Find
        $find.Text = $key
        $find.Replacement.Text = $Replacements[$key]
        $find.Forward = $true
        $find.Wrap = 1    # wdFindContinue
        $find.MatchCase = $false
        $find.MatchWholeWord = $false
        $find.MatchWildcards = $false
        $find.MatchSoundsLike = $false
        $find.MatchAllWordForms = $false
        
        # Replace all occurrences
        $null = $find.Execute(
            $key, $false, $false, $false, $false,
            $false, $true, 1, $true,
            $Replacements[$key], 2
        )
    }
    
    $doc.Save()
    $doc.Close()
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

# Main script
try {
    # Validate template paths
    if (-not (Test-Path -LiteralPath $basePaths.WelcomeLetterTemplate) -or
        -not (Test-Path -LiteralPath $basePaths.SecureLetterTemplate)) {
        throw "One or both document paths are invalid. Please check and try again."
    }

    # User input
    $rawFullName = Read-Host "Enter full name (First Last)"
    $startDate = Read-Host "Enter start date (MM/DD format)"

    # Process name
    $ti = (Get-Culture).TextInfo
    $fullName = $ti.ToTitleCase($rawFullName.ToLower())
    $parts = $fullName.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -lt 2) { throw "Please enter both first and last name." }
    
    # Generate credentials
    $username = ("{0}.{1}" -f $parts[0], $parts[-1]).ToLower()
    $initials = ($parts[0][0] + $parts[-1][0]).ToLower()
    $password = "BG-$initials$($startDate.Replace('/', ''))!@"

    # Show credentials
    Write-Host "Full Name: $fullName" -ForegroundColor Cyan
    Write-Host "Username: $username" -ForegroundColor Cyan
    Write-Host "Initial Password: $password" -ForegroundColor Cyan

    # Create working copies
    $tempDocs = @{
        WelcomeLetter = Join-Path (Split-Path $basePaths.WelcomeLetterTemplate) "temp_$(Split-Path $basePaths.WelcomeLetterTemplate -Leaf)"
        SecureLetter  = Join-Path (Split-Path $basePaths.SecureLetterTemplate) "temp_$(Split-Path $basePaths.SecureLetterTemplate -Leaf)"
    }
    Copy-Item -LiteralPath $basePaths.WelcomeLetterTemplate -Destination $tempDocs.WelcomeLetter -Force
    Copy-Item -LiteralPath $basePaths.SecureLetterTemplate -Destination $tempDocs.SecureLetter -Force

    # Replacement values
    $replacements = @{
        Welcome = @{
            "[[FULL_NAME]]" = $fullName
            "[[USERNAME]]"  = $username
        }
        Secure  = @{
            "[[FULL_NAME]]"  = $fullName
            "[[USERNAME]]"   = $username
            "[[PASSWORD]]"   = $password
            "[[START_DATE]]" = $startDate
        }
    }

    # Process documents
    Write-Host "Editing documents..." -ForegroundColor Yellow
    Set-WordText -DocPath $tempDocs.WelcomeLetter -Replacements $replacements.Welcome
    Set-WordText -DocPath $tempDocs.SecureLetter -Replacements $replacements.Secure

    # Create output directory
    $userFolder = Join-Path $basePaths.DesktopRoot $username
    if (-not (Test-Path -LiteralPath $userFolder)) {
        New-Item -ItemType Directory -Path $userFolder | Out-Null
    }

    # Generate PDF paths
    $outputPDFs = @{
        WelcomeLetter = Join-Path $userFolder ""
        SecureLetter  = Join-Path $userFolder ($basePaths.SecurePDFNameTemplate -f $username)
    }

    # Convert to PDF
    Write-Host "Converting to PDFs..." -ForegroundColor Yellow
    ConvertTo-PDF -WordPath $tempDocs.WelcomeLetter -PDFPath $outputPDFs.WelcomeLetter
    ConvertTo-PDF -WordPath $tempDocs.SecureLetter -PDFPath $outputPDFs.SecureLetter

    # Cleanup temp files
    Remove-Item -LiteralPath $tempDocs.WelcomeLetter -Force
    Remove-Item -LiteralPath $tempDocs.SecureLetter -Force

    Write-Host "Process completed successfully!" -ForegroundColor Green
    Write-Host "PDFs created:`n1. $($outputPDFs.WelcomeLetter)`n2. $($outputPDFs.SecureLetter)" -ForegroundColor Cyan
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
finally {
    # Open all relevant documents
    $pdfsToOpen = @(
        $outputPDFs.WelcomeLetter,
        $basePaths.SecureEmailGuide,
        $basePaths.AutoPilotGuide
        # (Get-ChildItem -Path $basePaths.DownloadsRoot -Filter "*.pdf" -ErrorAction SilentlyContinue |
        # Where-Object { $_.Name -match $basePaths.FedExPattern }).FullName
    )

    # Collect other relevant PDFs, including those matching the FedEx pattern
    $relevantPDFFiles = Get-ChildItem -Path $basePaths.DownloadsRoot -Filter "*.pdf" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match $FedExPattern -or $_.Name -match "OtherPattern1" -or $_.Name -match "OtherPattern2" } # Add more patterns as needed

    # Add each relevant PDF file path to the pdfsToOpen array
    foreach ($file in $relevantPDFFiles) {
        $pdfsToOpen += $file.FullName
    }

    foreach ($pdf in $pdfsToOpen) {
        if (Test-Path $pdf) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$pdf`""
        }
    }

    # Build email and copy to clipboard
    $email = "$username@"
    Set-Clipboard -Value $email

    # URLs
    $intuneUrl = "https://"
    $smartsheetUrl = "https:/"
    $emailUrl = "C:\Users"
    $teamsInviteUrl = "C:\Users"

    # Create form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Intune Group Access"
    $form.Size = New-Object System.Drawing.Size(500, 250)
    $form.StartPosition = "CenterScreen"

    # Label for email
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Your email (`"$email`") has been copied to the clipboard."
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($lbl)

    # Intune LinkLabel
    $intuneLink = New-Object System.Windows.Forms.LinkLabel
    $intuneLink.Text = "Intune GBL exception Group Link"
    $intuneLink.AutoSize = $true
    $intuneLink.Location = New-Object System.Drawing.Point(10, 60)
    $intuneLink.Links.Add(0, $intuneLink.Text.Length, $intuneUrl)
    $intuneLink.add_LinkClicked({
            param($control, $linkEvent)
            [System.Diagnostics.Process]::Start($linkEvent.Link.LinkData)
        })
    $form.Controls.Add($intuneLink)


    # Email Template LinkLabel
    $emailLink = New-Object System.Windows.Forms.LinkLabel
    $emailLink.Text = "Open Welcome Email Template, Attach secure email, and send to new hire"
    $emailLink.LinkColor = [System.Drawing.Color]::Blue
    $emailLink.AutoSize = $true
    $emailLink.Location = New-Object System.Drawing.Point(10, 90)
    $emailLink.Links.Add(0, $emailLink.Text.Length, $emailUrl)
    $emailLink.add_LinkClicked({
            param($control, $linkEvent)
            [System.Diagnostics.Process]::Start($linkEvent.Link.LinkData)
        })
    $form.Controls.Add($emailLink)

    # Email Template LinkLabel
    $teamsInvLink = New-Object System.Windows.Forms.LinkLabel
    $teamsInvLink.Text = "Optional New Hire IT 1:1 Introduction"
    $teamsInvLink.LinkColor = [System.Drawing.Color]::Blue
    $teamsInvLink.AutoSize = $true
    $teamsInvLink.Location = New-Object System.Drawing.Point(10, 120)
    $teamsInvLink.Links.Add(0, $teamsInvLink.Text.Length, $teamsInviteUrl)
    $teamsInvLink.add_LinkClicked({
            param($control, $linkEvent)
            [System.Diagnostics.Process]::Start($linkEvent.Link.LinkData)
        })
    $form.Controls.Add($teamsInvLink)

    # Smartsheet LinkLabel
    $ssLink = New-Object System.Windows.Forms.LinkLabel
    $ssLink.Text = "Update New Hire Smartsheet"
    $ssLink.AutoSize = $true
    $ssLink.Location = New-Object System.Drawing.Point(10, 150)
    $ssLink.Links.Add(0, $ssLink.Text.Length, $smartsheetUrl)
    $ssLink.add_LinkClicked({
            param($control, $linkEvent)
            [System.Diagnostics.Process]::Start($linkEvent.Link.LinkData)
        })
    $form.Controls.Add($ssLink)

    [void]$form.ShowDialog()

    # Cleanup Word processes
    Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
}
