# define script folder
$scriptFolder = $PSScriptRoot

# define functions 
$fFunction1 = join-Path $scriptFolder "openFolder.ps1"

#load powershell functions
."$fFunction1"

# select folder
#The starting folder to analyze
$startFolder = Get-SHDOpenFolderDialog -Title "Select the root folder for the Table of Contents"

Set-Location -Path $startFolder


# This script uses PowerShell modules and functions to convert various file types to PDF
function Convert-OfficeToPDF {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$DeleteOriginal = $false,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeSubfolders = $false,
        
        [Parameter(Mandatory = $false)]
        [string[]]$FileTypes = @('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.htm', '.html', '.txt', '.rtf', '.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff', '.bmp', '.csv')
    )
    
    # Set output path to input path if not specified
    if (-not $OutputPath) {
        $OutputPath = $FolderPath
    }
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory | Out-Null
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Green
    }
    
    # Get files to process
    $searchOptions = @{
        Path = $FolderPath
        Include = $FileTypes | ForEach-Object { "*$_" }
    }
    
    if ($IncludeSubfolders) {
        $searchOptions.Add('Recurse', $true)
    }
    
    $files = Get-ChildItem @searchOptions
    Write-Host "Found $($files.Count) files to process" -ForegroundColor Cyan
    
    # Initialize counters
    $successful = 0
    $failed = 0
    $skipped = 0
    
    # Create COM objects outside the loop for better performance
    try {
        # Initialize application objects
        $wordApp = $null
        $excelApp = $null 
        $powerpointApp = $null
        
        # Process files by type
        foreach ($file in $files) {
            Write-Host "Processing: $($file.FullName)" -ForegroundColor Yellow
            
            # Determine output filename
            $outputFile = Join-Path -Path $OutputPath -ChildPath "$($file.BaseName).pdf"
            
            # Skip if PDF already exists
            if (Test-Path -Path $outputFile) {
                Write-Host "  Skipping: PDF already exists at $outputFile" -ForegroundColor Gray
                $skipped++
                continue
            }
            
            $extension = $file.Extension.ToLower()
            $success = $false
            
            try {
                switch -Regex ($extension) {
                    # Word documents and text-based files
                    '^\.doc|\.docx|\.txt|\.rtf|\.htm|\.html|\.eml|\.odt$' {
                        if ($null -eq $wordApp) {
                            $wordApp = New-Object -ComObject Word.Application
                            $wordApp.Visible = $false
                            $wordApp.DisplayAlerts = $false
                        }
                        
                        $doc = $wordApp.Documents.Open($file.FullName)
                        $doc.SaveAs([ref]$outputFile, [ref]17) # 17 = PDF format
                        $doc.Close([ref]$false)
                        $success = $true
                        break
                    }
                    
                    # Excel files
                    '^\.xls|\.xlsx|\.csv$' {
                        if ($null -eq $excelApp) {
                            $excelApp = New-Object -ComObject Excel.Application
                            $excelApp.Visible = $false
                            $excelApp.DisplayAlerts = $false
                        }
                        
                        $workbook = $excelApp.Workbooks.Open($file.FullName)
                        $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $outputFile)
                        $workbook.Close($false)
                        $success = $true
                        break
                    }
                    
                    # PowerPoint files
                    '^\.ppt|\.pptx$' {
                        if ($null -eq $powerpointApp) {
                            $powerpointApp = New-Object -ComObject PowerPoint.Application
                            $powerpointApp.Visible = $false
                        }
                        
                        $presentation = $powerpointApp.Presentations.Open($file.FullName, $false, $false, $false)
                        $presentation.SaveAs($outputFile, 32) # 32 = ppSaveAsPDF
                        $presentation.Close()
                        $success = $true
                        break
                    }
                    
                    # Image files
                    '^\.jpg|\.jpeg|\.png|\.gif|\.tif|\.tiff|\.bmp$' {
                        if ($null -eq $wordApp) {
                            $wordApp = New-Object -ComObject Word.Application
                            $wordApp.Visible = $false
                            $wordApp.DisplayAlerts = $false
                        }
                        
                        $doc = $wordApp.Documents.Add()
                        $wordApp.Selection.InlineShapes.AddPicture($file.FullName) | Out-Null
                        $doc.SaveAs([ref]$outputFile, [ref]17) # 17 = PDF format
                        $doc.Close([ref]$false)
                        $success = $true
                        break
                    }
                    
                    default {
                        Write-Host "  Skipping: Unsupported file type - $extension" -ForegroundColor Gray
                        $skipped++
                    }
                }
                
                if ($success) {
                    Write-Host "  Success: Created $outputFile" -ForegroundColor Green
                    $successful++
                    
                    # Delete original if requested
                    if ($DeleteOriginal) {
                        Remove-Item -Path $file.FullName -Force
                        Write-Host "  Deleted original file: $($file.FullName)" -ForegroundColor DarkYellow
                    }
                }
            }
            catch {
                Write-Host "  Error processing $($file.FullName): $($_.Exception.Message)" -ForegroundColor Red
                $failed++
            }
            
            # Restart COM objects every 100 files to prevent memory issues
            if (($successful + $failed) % 100 -eq 0) {
                Write-Host "Restarting COM objects to prevent memory issues..." -ForegroundColor Cyan
                RestartComObjects
            }
        }
    }
    catch {
        Write-Host "Critical error: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # Clean up COM objects
        CleanupComObjects
        
        # Display summary
        Write-Host "`nConversion Summary:" -ForegroundColor Cyan
        Write-Host "  Successful: $successful" -ForegroundColor Green
        Write-Host "  Failed: $failed" -ForegroundColor Red
        Write-Host "  Skipped: $skipped" -ForegroundColor Gray
    }
    
    # Helper function to clean up COM objects
    function CleanupComObjects {
        if ($null -ne $wordApp) {
            try { $wordApp.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        }
        
        if ($null -ne $excelApp) {
            try { $excelApp.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        }
        
        if ($null -ne $powerpointApp) {
            try { $powerpointApp.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpointApp) | Out-Null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    
    # Helper function to restart COM objects
    function RestartComObjects {
        CleanupComObjects
        $wordApp = $excelApp = $powerpointApp = $null
    }
}

Convert-OfficeToPDF -FolderPath $startFolder


# Example usage:
# Convert-OfficeToPDF -FolderPath "E:\ExternalCertificatesWEBARCHIVEToPDF" -IncludeSubfolders
# Convert-OfficeToPDF -FolderPath "E:\Documents" -OutputPath "E:\PDFs" -DeleteOriginal -IncludeSubfolders
