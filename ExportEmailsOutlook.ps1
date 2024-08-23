# Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Outlook Email Exporter"
$form.Size = New-Object System.Drawing.Size(400, 400)
$form.StartPosition = "CenterScreen"

# Create a label for the shared mailbox dropdown
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(120, 20)
$label.Text = "Shared Mailbox:"
$form.Controls.Add($label)

# Create a ComboBox for selecting the shared mailbox
$mailboxComboBox = New-Object System.Windows.Forms.ComboBox
$mailboxComboBox.Location = New-Object System.Drawing.Point(130, 20)
$mailboxComboBox.Size = New-Object System.Drawing.Size(240, 20)
$mailboxComboBox.DropDownStyle = 'DropDownList'
$form.Controls.Add($mailboxComboBox)

# Create a button to browse folders
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(130, 50)
$browseButton.Size = New-Object System.Drawing.Size(100, 30)
$browseButton.Text = "Browse Folders"
$form.Controls.Add($browseButton)

# Create a TreeView to display folders and subfolders
$folderTreeView = New-Object System.Windows.Forms.TreeView
$folderTreeView.Location = New-Object System.Drawing.Point(10, 90)
$folderTreeView.Size = New-Object System.Drawing.Size(360, 180)
$folderTreeView.CheckBoxes = $true
$form.Controls.Add($folderTreeView)

# Create a label for the progress bar
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(10, 280)
$progressLabel.Size = New-Object System.Drawing.Size(360, 20)
$progressLabel.Text = "Progress:"
$form.Controls.Add($progressLabel)

# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 300)
$progressBar.Size = New-Object System.Drawing.Size(360, 30)
$form.Controls.Add($progressBar)

# Create a button to start the export
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(150, 340)
$exportButton.Size = New-Object System.Drawing.Size(100, 30)
$exportButton.Text = "Export Emails"
$form.Controls.Add($exportButton)

# Get Outlook Namespace
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Populate the ComboBox with shared mailboxes
foreach ($folder in $namespace.Folders) {
    $mailboxComboBox.Items.Add($folder.Name)
}

# Function to populate the TreeView with folders and subfolders
function Populate-TreeView ($parentNode, $folder) {
    foreach ($subFolder in $folder.Folders) {
        $node = $parentNode.Nodes.Add($subFolder.Name)
        Populate-TreeView -parentNode $node -folder $subFolder
    }
}

# Function to get the selected folders from the TreeView
function Get-SelectedFolders ($parentFolder, $node) {
    $selectedFolders = @()

    if ($node.Checked) {
        $folder = Get-OutlookFolder -folderName $node.Text -parentFolder $parentFolder
        if ($folder) {
            Write-Output "Selected Folder: $($folder.FolderPath)"  # Debugging
            $selectedFolders += $folder
        }
    }

    foreach ($childNode in $node.Nodes) {
        $selectedFolders += Get-SelectedFolders -parentFolder $parentFolder -node $childNode
    }

    return $selectedFolders
}

# Browse button click event
$browseButton.Add_Click({
    $selectedMailbox = $mailboxComboBox.SelectedItem
    if ($selectedMailbox) {
        $sharedFolder = $namespace.Folders.Item($selectedMailbox)
    } else {
        $sharedFolder = $namespace.Folders.Item(1)  # Default Inbox if no shared mailbox is selected
    }

    $folderTreeView.Nodes.Clear()
    foreach ($folder in $sharedFolder.Folders) {
        $node = $folderTreeView.Nodes.Add($folder.Name)
        Populate-TreeView -parentNode $node -folder $folder
    }
    $folderTreeView.ExpandAll()  # Expand all nodes by default
})

# Export button click event
$exportButton.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $saveDialog.ShowDialog()
    if ($result -eq "OK") {
        $saveDirectory = $saveDialog.SelectedPath
        $totalItems = 0
        $selectedFolders = @()

        # Collect all selected folders
        foreach ($node in $folderTreeView.Nodes) {
            $selectedFolders += Get-SelectedFolders -parentFolder $sharedFolder -node $node
        }

        # Verify selected folders
        Write-Output "Total folders selected: $($selectedFolders.Count)"

        # Calculate total number of items for the progress bar
        foreach ($folder in $selectedFolders) {
            $totalItems += $folder.Items.Count
        }
        $progressBar.Maximum = $totalItems
        $progressBar.Value = 0

        # Export emails from each selected folder
        foreach ($folder in $selectedFolders) {
            $folderPath = Join-Path $saveDirectory ($folder.FolderPath -replace "^\\", "" -replace "\\", "\")
            if (-not (Test-Path -Path $folderPath)) {
                Write-Output "Creating directory: $folderPath"  # Debugging log
                New-Item -ItemType Directory -Path $folderPath | Out-Null
            }

            foreach ($item in $folder.Items) {
                try {
                    if ($item.Class -eq 43) {  # MailItem
                        $subject = $item.Subject.Replace(":", "_").Replace("/", "_").Replace("\", "_").Replace("?", "_").Replace("*", "_").Replace("[", "_").Replace("]", "_")
                        $filePath = Join-Path $folderPath "$subject.msg"
                        Write-Output "Attempting to save email to: $filePath"  # Debugging log
                        $item.SaveAs($filePath, 3)  # 3 specifies the .msg format
                    }
                } catch {
                    Write-Error "Error saving email: $_"
                }
                $progressBar.Value += 1
                $form.Refresh()  # Refresh the form to update the progress bar
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Emails exported successfully!", "Success")
        $progressBar.Value = 0  # Reset the progress bar after completion
    }
})

# Function to get Outlook folder by name
function Get-OutlookFolder {
    param (
        [string]$folderName,
        [Microsoft.Office.Interop.Outlook.MAPIFolder]$parentFolder
    )
    foreach ($folder in $parentFolder.Folders) {
        if ($folder.Name -eq $folderName) {
            return $folder
        }
    }
    return $null
}

# Show the form
$form.ShowDialog()
