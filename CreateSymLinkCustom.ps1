Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Function to check for elevated privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to confirm overwrite
function Confirm-Overwrite {
    param ($linkPath)

	# Presume that it's not a folder
    $isDirectory = $false

    if (-not [string]::IsNullOrWhiteSpace($linkPath)) {
		# If the input box isn't empty, check if it's a valid path
        if (Test-Path $linkPath -PathType Leaf) {
			# if the path points to a file, it is not a folder
            $isDirectory = $false
        }
        elseif (Test-Path $linkPath -PathType Container) {
			# if the path points to a folder, it is a folder
            $isDirectory = $true
        }
    }

    $message = "A $(if ($isDirectory) { 'folder' } else { 'file' }) with the name '$(Split-Path $linkPath -Leaf)' already exists. Overwrite?"
    $result = [System.Windows.MessageBox]::Show(
        $message,
        "Confirm Overwrite",
        [System.Windows.MessageBoxButton]::YesNoCancel,
        [System.Windows.MessageBoxImage]::Warning
    )
    return $result
}

# Save the initial working directory
$initialWorkingDirectory = (Get-Location).Path

if (-not (Test-Admin)) {
    [System.Windows.MessageBox]::Show("Please run this script as an administrator. Restarting with elevated privileges...", "Elevation Required", "OK", "Warning")
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$initialWorkingDirectory`"" -Verb RunAs
    Exit
}

# Retrieve the working directory from arguments
$originalWorkingDirectory = $args[0]
Set-Location -Path $originalWorkingDirectory

# Create Window
$window = New-Object Windows.Window
$window.Title = "Symbolic Link Creator"
$window.Width = 400
$window.Height = 220
$window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
# Prevent resizing
$window.ResizeMode = [System.Windows.ResizeMode]::NoResize

# Create Grid Layout
$grid = New-Object Windows.Controls.Grid

# Create description label to explain what the textbox below is for
$inputLabel = New-Object System.Windows.Controls.Label
$inputLabel.Content = "Create Symbolic Link to:"
$inputLabel.VerticalAlignment = "Top"
$inputLabel.Margin = "10,0,10,10"
# Add label to the grid
$grid.Children.Add($inputLabel)

# Create TextBox for Path Input
$inputBox = New-Object Windows.Controls.TextBox
$inputBox.Margin = "10,25,10,10"
$inputBox.HorizontalAlignment = "Stretch"
$inputBox.VerticalAlignment = "Top"
# Add text box to the grid
$grid.Children.Add($inputBox)

# Create Buttons
$folderButton = New-Object Windows.Controls.Button
$folderButton.Content = "Browse Folder..."
$folderButton.Margin = "10,65,210,10"
$folderButton.Width = 150
$folderButton.HorizontalAlignment = "Left"
$folderButton.VerticalAlignment = "Top"

$fileButton = New-Object Windows.Controls.Button
$fileButton.Content = "Browse File..."
$fileButton.Margin = "210,65,10,10"
$fileButton.Width = 150
$fileButton.HorizontalAlignment = "Right"
$fileButton.VerticalAlignment = "Top"

$createButton = New-Object Windows.Controls.Button
$createButton.Content = "Create Sym Link!"
$createButton.Margin = "10,105,10,10"
$createButton.Width = 360
$createButton.HorizontalAlignment = "Stretch"
$createButton.VerticalAlignment = "Top"

$exitButton = New-Object Windows.Controls.Button
$exitButton.Content = "Exit"
$exitButton.Margin = "10,150,10,10"
$exitButton.Width = 360
$exitButton.HorizontalAlignment = "Stretch"
$exitButton.VerticalAlignment = "Top"

# Create the checkbox
$remainOpenAfterCreatingCheckBox = New-Object System.Windows.Controls.CheckBox
$remainOpenAfterCreatingCheckBox.Content = "Remain open after creation (to create multiple symlinks here)"
$remainOpenAfterCreatingCheckBox.IsChecked = $false
$remainOpenAfterCreatingCheckBox.VerticalAlignment = "Top"
$remainOpenAfterCreatingCheckBox.Margin = "10,130,10,10"

# Add controls to grid
$grid.Children.Add($folderButton)
$grid.Children.Add($fileButton)
$grid.Children.Add($createButton)
$grid.Children.Add($remainOpenAfterCreatingCheckBox)
$grid.Children.Add($exitButton)


# Folder button click handler
$folderButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select target folder"

    if (-not [string]::IsNullOrWhiteSpace($inputBox.Text) -and (Test-Path $inputBox.Text -PathType Container)) {
        $dialog.SelectedPath = $inputBox.Text
    }
    else {
        $dialog.SelectedPath = Get-Location
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputBox.Text = $dialog.SelectedPath
    }
})

# File button click handler
$fileButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select target file"

    if (-not [string]::IsNullOrWhiteSpace($inputBox.Text)) {
        if (Test-Path $inputBox.Text -PathType Leaf) {
            $dialog.InitialDirectory = Split-Path $inputBox.Text -Parent
            $dialog.FileName = Split-Path $inputBox.Text -Leaf
        }
        elseif (Test-Path $inputBox.Text -PathType Container) {
            $dialog.InitialDirectory = $inputBox.Text
        }
    }
    else {
        $dialog.InitialDirectory = Get-Location
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputBox.Text = $dialog.FileName
    }
})

# Create button click handler
$createButton.Add_Click({
    $targetPath = $inputBox.Text.Trim()
    $currentDir = Get-Location
    
    if (-not (Test-Path $targetPath)) {
        [System.Windows.MessageBox]::Show("Target path does not exist!", "Error", "OK", "Error")
        return
    }

    $linkName = Split-Path $targetPath -Leaf
    $linkPath = Join-Path -Path $currentDir -ChildPath $linkName

    try {
        if (Test-Path $linkPath) {
            $result = Confirm-Overwrite $linkPath
            if ($result -ne "Yes") {
                [System.Windows.MessageBox]::Show("Operation canceled", "Info", "OK", "Information")
                return
            }
            Remove-Item $linkPath -Force -Recurse -ErrorAction Stop
        }

        $isDirectory = (Get-Item $targetPath) -is [System.IO.DirectoryInfo]
        New-Item -ItemType SymbolicLink -Path $linkPath -Target $targetPath -ErrorAction Stop | Out-Null

        # Corrected message line
        $msg = "Successfully created symlink to $(if ($isDirectory) { 'folder' } else { 'file' })!"
		# Add the rest of the debug output ...
		$msg += "`n`n`n`n"
		$msg += "Original Working Directory:"
		$msg += "`n"
		$msg += "$originalWorkingDirectory"
		$msg += "`n`n"
		$msg += "Current Directory:"
		$msg += "`n"
		$msg += "$currentDir"
		$msg += "`n`n"
		$msg += "Sym Link Target (point-to) Path:"
		$msg += "`n"
		$msg += "$targetPath"
		$msg += "`n`n"
		$msg += "Sym Link File (destination) Path:"
		$msg += "`n"
		$msg += "$linkPath"
		
        [System.Windows.MessageBox]::Show($msg, "Success", "OK", "Information")
    }
    catch {
        [System.Windows.MessageBox]::Show("Error creating symlink: $($_.Exception.Message)", "Error", "OK", "Error")
    }
	
	if ($remainOpenAfterCreatingCheckBox.IsChecked -eq $true) {
		# Do nothing, remain open
	} else {
		$window.Close()
	}
})

# Exit button click handler
$exitButton.Add_Click({
	$window.Close()
})

# Set and show Window content
$window.Content = $grid
$window.ShowDialog() | Out-Null