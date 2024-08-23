# Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Outlook Email Exporter"
$form.Size = New-Object System.Drawing.Size(400, 400)
$form.StartPosition = "CenterScreen"

# Create a ComboBox for selecting the shared mailbox
$mailboxComboBox = New-Object System.Windows.Forms.ComboBox
$mailboxComboBox.Location = New-Object System.Drawing.Point(130, 20)
$mailboxComboBox.Size = New-Object System.Drawing.Size(240, 20)
$mailboxComboBox.DropDownStyle = 'DropDownList'
$form.Controls.Add($mailboxComboBox)

# Create a TreeView to display folders and subfolders
$folderTreeView = New-Object System.Windows.Forms.TreeView
$folderTreeView.Location = New-Object System.Drawing.Point(10, 50)
$folderTreeView.Size = New-Object System.Drawing.Size(360, 200)
$folderTreeView.CheckBoxes = $true
$form.Controls.Add($folderTreeView)

# Create a button to start the export
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(150, 270)
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

# Browse button click event
$mailboxComboBox.Add_SelectedIndexChanged({
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
    $selectedFolders = @()

    # Collect all selected folders
    foreach ($node in $folderTreeView.Nodes) {
        if ($node.Checked) {
            $folder = Get-OutlookFolder -folderName $node.Text -parentFolder $sharedFolder
            if ($folder) {
                $selectedFolders += $folder
                Write-Output "Selected Folder: $($folder.FolderPath)"
            }
        }
    }

    foreach ($folder in $selectedFolders) {
        Write-Output "Processing folder: $($folder.Name) with $($folder.Items.Count) items"
        foreach ($item in $folder.Items) {
            if ($item.Class -eq 43) {  # MailItem
                $subject = $item.Subject.Replace(":", "_").Replace("/", "_").Replace("\", "_").Replace("?", "_").Replace("*", "_").Replace("[", "_").Replace("]", "_")
                $filePath = Join-Path "C:\Temp\ExportedEmails" "$subject.msg"
                Write-Output "Saving email: $subject"
                $item.SaveAs($filePath, 3)  # 3 specifies the .msg format
            }
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Emails exported successfully!", "Success")
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
