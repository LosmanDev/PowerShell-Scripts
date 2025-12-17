# -----------------------------------------
# 2) Policy Manager Section (CSP Applied)
# -----------------------------------------

$PolicyRoot = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device"
$PolicyList = [System.Collections.Generic.List[Object]]::new()

if (Test-Path $PolicyRoot) {

    Get-ChildItem -Recurse $PolicyRoot | ForEach-Object {

        $path = $_.PSPath
        $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue

        foreach ($p in $props.PSObject.Properties) {

            if ($p.Name -in "PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider") { continue }

            $category = ($path -replace '^.*device\\', '') -replace '\\.*', ''

            $PolicyList.Add([PSCustomObject]@{
                    Source      = "PolicyManager"
                    Category    = $category
                    Name        = $p.Name
                    Version     = $null
                    InstallDate = $null
                    Value       = $p.Value
                    RegistryKey = $path
                })
        }
    }
}

$PolicyList |
Sort-Object Source, Category, Name |
Out-GridView -Title "Policy Manager View"

