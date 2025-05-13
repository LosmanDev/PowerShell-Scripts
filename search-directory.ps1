# Define the path to the directory and the keywords to search for
$directoryPath = ""
$keywords = @("keyword", "test")

# Get all files in the directory
$files = Get-ChildItem -Path $directoryPath -Recurse

# Initialize progress bar
$totalFiles = $files.Count
$currentFile = 0

# Arrays to store results
$filesWithKeywordInName = @()
$filesWithKeywordInContent = @()

foreach ($file in $files) {
    # Update progress bar
    $currentFile++
    Write-Progress -Activity "Searching files" -Status "Processing $currentFile of $totalFiles" -PercentComplete (($currentFile / $totalFiles) * 100)
    
    if ($file.PSIsContainer -eq $false) {
        # Diagnostic output
        Write-Output "Processing file: $($file.FullName)"
        
        foreach ($keyword in $keywords) {
            # Search for the keyword in the file name (case-insensitive)
            if ($file.Name -ilike "*$keyword*") {
                Write-Output "Keyword '$keyword' found in file name: $($file.FullName)"
                $filesWithKeywordInName += $file.FullName
            }

            # Search for the keyword in the file content (case-insensitive)
            if (Select-String -Path $file.FullName -Pattern $keyword -Quiet -CaseSensitive:$false) {
                Write-Output "Keyword '$keyword' found in file content: $($file.FullName)"
                $filesWithKeywordInContent += $file.FullName
            }
        }
    }
}

# Combine results
$allFilesWithKeyword = $filesWithKeywordInName + $filesWithKeywordInContent

# Remove duplicates
$uniqueFilesWithKeyword = $allFilesWithKeyword | Select-Object -Unique

# Output the results
if ($uniqueFilesWithKeyword.Count -eq 0) {
    Write-Output "No files found with the keywords."
}
else {
    $uniqueFilesWithKeyword | ForEach-Object { Write-Output $_ }
}
