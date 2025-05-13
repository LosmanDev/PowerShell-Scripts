# Load required assemblies
Add-Type -AssemblyName Microsoft.Office.Interop.Word

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
    # — Fixed document paths
    $doc1Path = ""
    
    #Secure Letter
    $doc2Path = ""
  
    # — Validate literal paths (handles the [name] bracket)
    if (
        -not (Test-Path -LiteralPath $doc1Path) -or
        -not (Test-Path -LiteralPath $doc2Path)
    ) {
        throw "One or both document paths are invalid. Please check and try again."
    }


    # 3) Ask user
    $rawFullName = Read-Host "Enter full name (First Last)"
    $startDate = Read-Host "Enter start date (MM/DD format)"

    # 4) Title-case the name
    $ti = (Get-Culture).TextInfo
    $fullName = $ti.ToTitleCase($rawFullName.ToLower())

    # 5) Split into first/last, build lowercase username
    $parts = $fullName.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -lt 2) { throw "Please enter both first and last name." }
    $username = ("{0}.{1}" -f $parts[0], $parts[-1]).ToLower()

    # 6) Build the single password BG-xxMMDD!@
    $initials = ($parts[0][0] + $parts[-1][0]).ToLower()
    $dateFormatted = $startDate.Replace('/', '')
    $password = "BG-$initials$dateFormatted!@"

    # 7) Show on console
    Write-Host "Full Name: $fullName"                     -ForegroundColor Cyan
    Write-Host "Username: $username"                       -ForegroundColor Cyan
    Write-Host "Initial Password: $password"               -ForegroundColor Cyan

   

    # 8) Copy to temp docs
    $tempDoc1 = Join-Path (Split-Path $doc1Path) "temp_$(Split-Path $doc1Path -Leaf)"
    $tempDoc2 = Join-Path (Split-Path $doc2Path) "temp_$(Split-Path $doc2Path -Leaf)"

    Copy-Item -LiteralPath $doc1Path -Destination $tempDoc1 -Force
    Copy-Item -LiteralPath $doc2Path -Destination $tempDoc2 -Force

    # 9) Build replacement tables
    $repl1 = @{
        "[[FULL_NAME]]" = $fullName
        "[[USERNAME]]"  = $username
    }
    $repl2 = @{
        "[[FULL_NAME]]"  = $fullName
        "[[USERNAME]]"   = $username
        "[[PASSWORD]]"   = $password
        "[[START_DATE]]" = $startDate
    }
    
    # — Perform find&replace
    Write-Host "Editing documents..." -ForegroundColor Yellow
    Set-WordText -DocPath $tempDoc1 -Replacements $repl1
    Set-WordText -DocPath $tempDoc2 -Replacements $repl2
    
    # — Define output PDF names
    $pdf1Name = ""
    $pdf2Name = "$username"
  

    # — Build a desktop subfolder named after the user
    $desktop = 'C:\Users\'
    $userFolder = Join-Path $desktop $username

    if (-not (Test-Path -LiteralPath $userFolder)) {
        New-Item -ItemType Directory -Path $userFolder | Out-Null
    }

    # — Now put the PDFs into that new folder
    $pdf1Path = Join-Path $userFolder $pdf1Name
    $pdf2Path = Join-Path $userFolder $pdf2Name

    

    # — Convert to PDF
    Write-Host "Converting to PDFs..." -ForegroundColor Yellow
    ConvertTo-PDF -WordPath $tempDoc1 -PDFPath $pdf1Path
    ConvertTo-PDF -WordPath $tempDoc2 -PDFPath $pdf2Path
    
    # — Cleanup temp files (literal)
    Remove-Item -LiteralPath $tempDoc1 -Force
    Remove-Item -LiteralPath $tempDoc2 -Force
    
    Write-Host "Process completed successfully!" -ForegroundColor Green
    Write-Host "PDFs created:"                                     -ForegroundColor Cyan
    Write-Host "1. $pdf1Path"                                     -ForegroundColor Cyan
    Write-Host "2. $pdf2Path"                                     -ForegroundColor Cyan

     
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
finally {
    # Ensure Word quits if something went wrong
    Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    ForEach-Object { $_.Kill() } 
}

