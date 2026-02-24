# ComData Commercial POS “CWS_EOD_Update” File Checker
#COPYRIGHT © 2026 CIRCLE K STORES AND ALIMENTATION COUCHE-TARD.
Add-Type -AssemblyName PresentationFramework

# Create Window
$Window = New-Object System.Windows.Window
$Window.Title = "Comdata Commercial POS CWS_EOD_Update File Checker"
$Window.Width = 500
$Window.Height = 350
$Window.WindowStartupLocation = "CenterScreen"

# Create Grid
$Grid = New-Object System.Windows.Controls.Grid
$Window.Content = $Grid
1..4 | ForEach-Object { $Grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition)) }

# Store Number Label
$Label = New-Object System.Windows.Controls.Label
$Label.Content = "Enter Store Number(s):"
$Label.HorizontalAlignment = "Center"
$Label.VerticalAlignment = "Center"
[System.Windows.Controls.Grid]::SetRow($Label, 0)
$Grid.Children.Add($Label)

# Store TextBox
$TextBox = New-Object System.Windows.Controls.TextBox
$TextBox.Width = 250
$TextBox.Height = 25
$TextBox.HorizontalAlignment = "Center"
$TextBox.VerticalAlignment = "Center"
[System.Windows.Controls.Grid]::SetRow($TextBox, 1)
$Grid.Children.Add($TextBox)

# TrendAR Folder Button
$FolderButton = New-Object System.Windows.Controls.Button
$FolderButton.Content = "Get The 10 Recently Modified Trendar Files"
$FolderButton.Width = 300
$FolderButton.Height = 30
$FolderButton.HorizontalAlignment = "Center"
$FolderButton.VerticalAlignment = "Center"
[System.Windows.Controls.Grid]::SetRow($FolderButton, 2)
$Grid.Children.Add($FolderButton)

# Button Click Event
$FolderButton.Add_Click({

    $inputStores = $TextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($inputStores)) {
        [System.Windows.MessageBox]::Show("Please enter at least one store number.")
        return
    }

    $storeList = $inputStores -split '[,;\s]+' | Where-Object { $_ -ne "" }

    # Prompt for credentials once
    $Cred = Get-Credential -Message "Enter domain credentials for all stores"
    if (-not $Cred) { return }

    # Get folder where script is located
    $scriptFolder = Split-Path -Parent $PSCommandPath
    if (-not $scriptFolder) { $scriptFolder = Get-Location }
    $csvPath = Join-Path -Path $scriptFolder -ChildPath "Top10RecentFiles_Master.csv"
    $writeHeaders = -not (Test-Path $csvPath)

    foreach ($StoreNumber in $storeList) {
        try {
            $ComputerName = "${StoreNumber}FPOS1"

            # Connect to remote store
            $session = New-PSSession -ComputerName $ComputerName -Credential $Cred -ErrorAction Stop

            # Get top 10 files
            $fileList = Invoke-Command -Session $session -ScriptBlock {
                $folder = 'C:\trendar\cws'
                if (Test-Path $folder) {
                    Get-ChildItem -Path $folder -File |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 10 |
                        Select-Object Name, LastWriteTime
                } else { @() }
            }

            Remove-PSSession $session

            if (-not $fileList -or $fileList.Count -eq 0) { continue }

            # Add StoreNumber
            $fileListWithStore = $fileList | ForEach-Object {
                [PSCustomObject]@{
                    StoreNumber  = $StoreNumber
                    FileName     = $_.Name
                    LastModified = $_.LastWriteTime
                }
            }

            # Append to CSV
            $fileListWithStore | Export-Csv -Path $csvPath -NoTypeInformation -Append:$(! $writeHeaders)
            $writeHeaders = $false

        } catch {
            [System.Windows.MessageBox]::Show("Failed for store $StoreNumber{}")
        }
    }

    [System.Windows.MessageBox]::Show("Top 10 recent files for all stores appended to CSV:`n$csvPath")
})

$Window.ShowDialog()
