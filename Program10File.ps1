# FPOS1 Remote Connector GUI - Grabs the Last 10 files in a file location and prints the file name and last modified date in a csv file for data
#COPYRIGHT © 2026 CIRCLE K STORES AND ALIMENTATION COUCHE-TARD.
Add-Type -AssemblyName PresentationFramework

# Create Window
$Window = New-Object System.Windows.Window
$Window.Title = "FPOS1 Last 10 Modified Files Info Grabber"
$Window.Width = 500
$Window.Height = 400
$Window.WindowStartupLocation = "CenterScreen"

# Create Grid
$Grid = New-Object System.Windows.Controls.Grid
$Window.Content = $Grid
1..5 | ForEach-Object { $Grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition)) }

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
$FolderButton.Content = "Get Last 10 Modified Files"
$FolderButton.Width = 300
$FolderButton.Height = 30
$FolderButton.HorizontalAlignment = "Center"
$FolderButton.VerticalAlignment = "Center"
[System.Windows.Controls.Grid]::SetRow($FolderButton, 2)
$Grid.Children.Add($FolderButton)

# Status TextBox
$StatusBox = New-Object System.Windows.Controls.TextBox
$StatusBox.Width = 350
$StatusBox.Height = 50
$StatusBox.HorizontalAlignment = "Center"
$StatusBox.VerticalAlignment = "Center"
$StatusBox.IsReadOnly = $true
$StatusBox.TextWrapping = "Wrap"
[System.Windows.Controls.Grid]::SetRow($StatusBox, 3)
$Grid.Children.Add($StatusBox)

# TrendAR Button Click Event
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

    # Get script folder for CSV
    $scriptFolder = Split-Path -Parent $PSCommandPath
    if (-not $scriptFolder) { $scriptFolder = Get-Location }
    $csvPath = Join-Path -Path $scriptFolder -ChildPath "Top10RecentFiles_Master.csv"

    # Overwrite if file exists
    if (Test-Path $csvPath) { Remove-Item $csvPath -Force }
    $writeHeaders = $true

    foreach ($StoreNumber in $storeList) {
        try {
            # Update status for user
            $StatusBox.Text = "Processing store $StoreNumber..."
            $StatusBox.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

            $ComputerName = "${StoreNumber}FPOS1"
            $session = New-PSSession -ComputerName $ComputerName -Credential $Cred -ErrorAction Stop

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

            $fileListWithStore = $fileList | ForEach-Object {
                [PSCustomObject]@{
                    StoreNumber  = $StoreNumber
                    FileName     = $_.Name
                    LastModified = $_.LastWriteTime
                }
            }

            # Export to CSV (overwrite already handled)
            $fileListWithStore | Export-Csv -Path $csvPath -NoTypeInformation -Append:$(! $writeHeaders)
            $writeHeaders = $false

        } catch {
            [System.Windows.MessageBox]::Show("Failed for store $StoreNumber{}")
        }
    }

    $StatusBox.Text = "Completed! Top 10 recent files saved to CSV:`n$csvPath"
})

$Window.ShowDialog()
