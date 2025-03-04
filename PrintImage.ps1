# Load required .NET types for printing and image handling
Add-Type -AssemblyName System.Drawing

# Specify your printer name and image file path.
$printerName = "Canon G5000 Series"
$imageFile = "C:\Users\Jordan\Documents\Color-Test-Page-by-PrintTester.jpg"

# Create a new PrintDocument object and set its printer
$printDoc = New-Object System.Drawing.Printing.PrintDocument
$printDoc.PrinterSettings.PrinterName = $printerName

# Check if the printer is valid
if (-not $printDoc.PrinterSettings.IsValid) {
    Write-Host "Printer '$printerName' is not valid. Check the printer name."
    exit
}

# Find the first paper source that indicates a manual or rear feed.
$manualSource = $printDoc.PrinterSettings.PaperSources |
    Where-Object { $_.SourceName -match "Manual" -or $_.SourceName -match "Rear" } |
    Select-Object -First 1

if ($manualSource) {
    Write-Host "Using paper source: $($manualSource.SourceName)"
    $printDoc.DefaultPageSettings.PaperSource = $manualSource
} else {
    Write-Host "Manual/rear feed not found. Please verify your printer settings."
    # Optionally exit or continue using the default feed
    # exit
}

# Load the image from the file
$image = [System.Drawing.Image]::FromFile($imageFile)

# Define the PrintPage event handler to render the image
$printDoc.add_PrintPage({
    param($sender, $e)
    # Draw the image scaled to the margin bounds
    $e.Graphics.DrawImage($image, $e.MarginBounds)
    $e.HasMorePages = $false
})

# Send the print job
$printDoc.Print()

# Clean up the image resource
$image.Dispose()
