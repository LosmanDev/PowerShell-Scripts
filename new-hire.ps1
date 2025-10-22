# Load required assemblies
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

function Show-Notification {
    param (
        [Parameter(Position = 0)]
        [string]$Message,
        [Parameter(Position = 1)]
        [string]$Title = "company company"
    )

    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = "company company" }

    $balloon = New-Object System.Windows.Forms.NotifyIcon
    $balloon.Icon = [System.Drawing.SystemIcons]::Information
    $balloon.Visible = $true
    try {
        $balloon.ShowBalloonTip(5000, $Title, $Message, [System.Windows.Forms.ToolTipIcon]::Info)
        Start-Sleep -Milliseconds 600
    }
    finally {
        $balloon.Dispose()
    }
}


$actualUsername = $MyInvocation.MyCommand.Path.Split('\')[2]
$actualUserProfile = "C:\Users\$actualUsername"

$basePaths = @{
    WelcomeLetterTemplate = Join-Path $actualUserProfile "OneDrive - company\Desktop\Onboarding\Welcome Letters\company company U.S. Welcome Letter Template.docx"
    SecureLetterTemplate  = Join-Path $actualUserProfile "OneDrive - company\Desktop\Onboarding\Welcome Letters\[name] Secure company company U.S. Welcome Letter.docx"
    
    DesktopRoot           = Join-Path $actualUserProfile "OneDrive - company\Desktop\New Hire Folders"
    DownloadsRoot         = Join-Path $actualUserProfile "Downloads"
    
    SecureEmailGuide      = Join-Path $actualUserProfile "OneDrive - company\Desktop\Onboarding\company company U.S. Secure Email - Accessing a Secure Email.pdf" 
    AutoPilotGuide        = Join-Path $actualUserProfile "OneDrive - company\Desktop\Onboarding\company company U.S. AutoPilot Laptop Configuration Instructions.pdf"
    EmailTemplate         = Join-Path $actualUserProfile "OneDrive - company\Desktop\Onboarding\Emails\Secure Welcome to company company!.msg"
    TeamsTemplate         = Join-Path $actualUserProfile "AppData\Roaming\Microsoft\Templates\(optional) New Hire IT 11 Introduction.oft"
    TeamsTemplateSource   = Join-Path $actualUserProfile "OneDrive - company\Desktop\Onboarding\Emails\(optional) New Hire IT 11 Introduction.oft"
    WhiteGloveTemplate    = Join-Path $actualUserProfile "OneDrive - company\Desktop\Onboarding\Emails\White Glove.msg"
    
    IntuneGroupURL        = ""
    SmartsheetURL         = ""
    
    SecurePDFNameTemplate = "{0} Secure company company U.S. Welcome Letter.pdf"
    FedExPattern          = "FedEx-Shipping-Label"
}

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

    function Show-InputDialog {
        param (
            [string]$Message,
            [string]$Title,
            [string]$DefaultValue = ""
        )
    
        $form = New-Object Windows.Forms.Form
        $form.Text = $Title
        $form.Size = New-Object Drawing.Size @(400, 200)
        $form.StartPosition = "CenterScreen"
        $form.TopMost = $true
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
    
        $label = New-Object Windows.Forms.Label
        $label.Location = New-Object Drawing.Point @(10, 20)
        $label.Size = New-Object Drawing.Size @(360, 40)
        $label.Text = $Message
        $form.Controls.Add($label)
    
        $textBox = New-Object Windows.Forms.TextBox
        $textBox.Location = New-Object Drawing.Point @(10, 70)
        $textBox.Size = New-Object Drawing.Size @(360, 20)
        $textBox.Text = $DefaultValue
        $form.Controls.Add($textBox)
    
        $okButton = New-Object Windows.Forms.Button
        $okButton.Location = New-Object Drawing.Point @(120, 120)
        $okButton.Size = New-Object Drawing.Size @(75, 23)
        $okButton.Text = "OK"
        $okButton.DialogResult = [Windows.Forms.DialogResult]::OK
        $form.Controls.Add($okButton)
        $form.AcceptButton = $okButton
    
        $cancelButton = New-Object Windows.Forms.Button
        $cancelButton.Location = New-Object Drawing.Point @(205, 120)
        $cancelButton.Size = New-Object Drawing.Size @(75, 23)
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = [Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($cancelButton)
        $form.CancelButton = $cancelButton
    
        $form.Activate()
        $result = $form.ShowDialog()
    
        if ($result -eq [Windows.Forms.DialogResult]::OK) {
            return $textBox.Text
        }
        return $null
    }
    
    $rawFullName = Show-InputDialog -Message "Enter full name (First Last, john doe, John Doe)`nUsername might differ, edit in PDF if that's the case" -Title "Full Name"
    if ([string]::IsNullOrEmpty($rawFullName)) {
        throw "Name input cancelled"
    }
    
    $startDate = Show-InputDialog -Message "Enter start date (MM/DD formats to use 0401, 04/01)" -Title "Start Date"
    if ([string]::IsNullOrEmpty($startDate)) {
        throw "Date input cancelled"
    }


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
        WelcomeLetter = Join-Path $userFolder "company company U.S. Welcome Letter.pdf"
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
    $email = "$username@company.com"
    Set-Clipboard -Value $email

    
    # Create form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "company company LTD New Hire Quick Links"
    $form.Size = New-Object System.Drawing.Size(500, 250)
    $form.StartPosition = "CenterScreen"

    # Label for email
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Your company email (`"$email`") has been copied to the clipboard."
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($lbl)

    # Intune LinkLabel
    $intuneLink = New-Object System.Windows.Forms.LinkLabel
    $intuneLink.Text = "Intune GBL exception Group Link"
    $intuneLink.AutoSize = $true
    $intuneLink.Location = New-Object System.Drawing.Point(10, 60)
    $intuneLink.Links.Add(0, $intuneLink.Text.Length, $basePaths.IntuneGroupURL)
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
    $emailLink.Links.Add(0, $emailLink.Text.Length, $basePaths.EmailTemplate)
    $emailLink.add_LinkClicked({
            param($control, $linkEvent)
            [System.Diagnostics.Process]::Start($linkEvent.Link.LinkData)
        })
    $form.Controls.Add($emailLink)

    $teamsInvLink = New-Object System.Windows.Forms.LinkLabel
    $teamsInvLink.Text = "Optional New Hire IT 1:1 Introduction [Add teams invite / Signature]"
    $teamsInvLink.LinkColor = [System.Drawing.Color]::Blue
    $teamsInvLink.AutoSize = $true
    $teamsInvLink.Location = New-Object System.Drawing.Point(10, 120)
    $teamsInvLink.Links.Add(0, $teamsInvLink.Text.Length, $basePaths.TeamsTemplate)
    
    $teamsInvLink.add_LinkClicked({
            param($control, $linkEvent)
    
            # Use dynamic user profile paths already in $basePaths.
            $src = $basePaths.TeamsTemplateSource
            $dest = $basePaths.TeamsTemplate
            try {
                if (-not (Test-Path -LiteralPath $src)) {
                    [System.Windows.Forms.MessageBox]::Show("Source template not found:`n$src", "Template missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }
    
                $destDir = Split-Path -Path $dest -Parent
                if (-not (Test-Path -LiteralPath $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }
    
                Copy-Item -LiteralPath $src -Destination $dest -Force
    
                # Open the copied .oft from the destination folder
                Start-Process -FilePath $dest
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to copy/open template:`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
    $form.Controls.Add($teamsInvLink)

    # Email Template LinkLabel
    $whiteGloveLink = New-Object System.Windows.Forms.LinkLabel
    $whiteGloveLink.Text = "White Glove Email Template"
    $whiteGloveLink.LinkColor = [System.Drawing.Color]::Blue
    $whiteGloveLink.AutoSize = $true
    $whiteGloveLink.Location = New-Object System.Drawing.Point(10, 150)
    $whiteGloveLink.Links.Add(0, $whiteGloveLink.Text.Length, $basePaths.WhiteGloveTemplate)
    $whiteGloveLink.add_LinkClicked({
            param($control, $linkEvent)
            [System.Diagnostics.Process]::Start($linkEvent.Link.LinkData)
        })
    $form.Controls.Add($whiteGloveLink)

    # Smartsheet LinkLabel
    $ssLink = New-Object System.Windows.Forms.LinkLabel
    $ssLink.Text = "Update New Hire Smartsheet"
    $ssLink.AutoSize = $true
    $ssLink.Location = New-Object System.Drawing.Point(10, 180)
    $ssLink.Links.Add(0, $ssLink.Text.Length, $basePaths.SmartsheetURL)
    $ssLink.add_LinkClicked({
            param($control, $linkEvent)
            [System.Diagnostics.Process]::Start($linkEvent.Link.LinkData)
        })
    $form.Controls.Add($ssLink)

    [void]$form.ShowDialog()

    $notif = "New Hire folder created:`nUsername: $username`nFolder: $userFolder`nFiles:`n - $($outputPDFs.WelcomeLetter)`n - $($outputPDFs.SecureLetter)"
    Show-Notification $notif

    # Cleanup Word processes
    Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
}
