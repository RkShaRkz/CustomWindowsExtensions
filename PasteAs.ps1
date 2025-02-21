# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to show "no file in clipboard" error
function Show-NoFileToCopyErrorDialog {
    [System.Windows.Forms.MessageBox]::Show(
        "There is no file in the clipboard to copy!",
        "Clipboard Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

# Function to show the "overwrite with self" dialog
function Show-OverwriteWithSelfDialog {
    $dialogResult = [System.Windows.Forms.MessageBox]::Show(
        "You are trying to overwrite the original file with itself. Are you sure you want to do this?",
        "Confirm Overwrite",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    return $dialogResult
}

# Function to show the overwrite dialog
function Show-OverwriteDialog {
    $dialogResult = [System.Windows.Forms.MessageBox]::Show(
        "The file already exists. Do you want to overwrite it?",
        "Confirm Overwrite",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    return $dialogResult
}

# Check if clipboard contains file drop list
try {
    $fileList = Get-Clipboard -Format FileDropList
}
catch {
    $fileList = @()
}

if ($fileList.Count -gt 0) {
    # define some handy shortcuts to more easily tweak the UI
    $checkboxX = 85
    $checkboxY = 30
    $buttonsY = 55

    # Create a Windows Forms input box
    $inputBox = New-Object System.Windows.Forms.Form
    $inputBox.Width = 320
    $inputBox.Height = 130
    $inputBox.Text = "Enter new filename"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "New filename:"
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $inputBox.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = (Get-Item $fileList[0]).BaseName
    $textBox.Location = New-Object System.Drawing.Point(120, 10)
    $textBox.Width = 180
    $inputBox.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(10, $buttonsY)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputBox.AcceptButton = $okButton  # Enter key triggers OK button
    $inputBox.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(220, $buttonsY)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inputBox.CancelButton = $cancelButton  # Escape key triggers Cancel button
    $inputBox.Controls.Add($cancelButton)

    $keepExtensionCheckBox = New-Object System.Windows.Forms.CheckBox
    $keepExtensionCheckBox.Text = "Keep original extension"
    $keepExtensionCheckBox.Location = New-Object System.Drawing.Point($checkboxX, $checkboxY)
    $keepExtensionCheckBox.AutoSize = $true  # Automatically size to content
    $keepExtensionCheckBox.Checked = $true   # Default checked state
    $inputBox.Controls.Add($keepExtensionCheckBox)

    # Run PowerShell script without showing the console window
    Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File $($MyInvocation.MyCommand.Path)" -WindowStyle Hidden

    # Show the input box and get the result
    $result = $inputBox.ShowDialog()

    # Process the result
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $newFile = $textBox.Text
        $originalFile = $fileList[0]
        $originalExtension = (Get-Item $originalFile).Extension

        if ($newFile -ne "") {
            # Get the folder path of the original file
            $folderPath = Split-Path -Path $originalFile -Parent

            # Build the full paths for source and destination files
            # If "keep extension" is checked, use the thing from the text box as just the filename, otherwise use it as the whole filename (with or without extension)
            if ($keepExtensionCheckBox.Checked) {
                $destinationFile = Join-Path -Path $folderPath -ChildPath "$newFile$originalExtension"
            } else {
                $destinationFile = Join-Path -Path $folderPath -ChildPath $newFile
            }

            # Copy the file to the new location
            # Check if the destination file exists
            if (Test-Path -Path $destinationFile) {
                # if the file already exists, normalize both so we can see if we're trying to overwrite the original file or some other file
                # Normalize paths for comparison
                $normalizedOriginalFile = (Get-Item -Path $originalFile).FullName
                $normalizedDestinationFile = (Get-Item -Path $destinationFile).FullName
                # now check what we're overwriting
                if ($normalizedOriginalFile -eq $normalizedDestinationFile) {
                    Write-Host "Warning: Trying to overwrite the file with itself."
                    $response = Show-OverwriteWithSelfDialog
                    if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Copy-Item -Path $file -Destination $destination -Force
                        Write-Host "File has been overwritten."
                    } else {
                        Write-Host "File was not overwritten."
                    }
                } else {
                    $response = Show-OverwriteDialog
                    if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Copy-Item -Path $file -Destination $destination -Force
                        Write-Host "File has been overwritten."
                    } else {
                        Write-Host "File was not overwritten."
                    }
                }
            } else {
                Copy-Item -Path $originalFile -Destination $destinationFile
                Write-Host "File has been copied successfully!"
            }

            Write-Host "File copied successfully!"
        }
        else {
            Write-Host "Error: The new name cannot be empty!"
        }
    }
}
else {
    Write-Host "Error: The clipboard does not contain a file drop list."
    Show-NoFileToCopyErrorDialog
}
