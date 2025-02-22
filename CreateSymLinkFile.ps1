# Function to check for elevated privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Save the initial working directory
$initialWorkingDirectory = (Get-Location).Path

if (-not (Test-Admin)) {
    Write-Output "Please run this script as an administrator. Let's try rerunning automatically..."
    
    # Re-run the script with elevated privileges, passing the initial working directory as an argument
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$initialWorkingDirectory`"" -Verb RunAs
    Exit  # Exit the non-elevated instance
}

# Retrieve the working directory from arguments
$originalWorkingDirectory = $args[0]
Set-Location -Path $originalWorkingDirectory

# Function to show file picker dialog
function Show-FilePickerDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = [System.IO.Directory]::GetCurrentDirectory()
    $dialog.Filter = "All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        return $null
    }
}

# Main script
$selectedPath = Show-FilePickerDialog

if ($selectedPath) {
    $currentDir = Get-Location -PSProvider FileSystem
    $linkName = Split-Path -Leaf $selectedPath
    $linkPath = Join-Path -Path $currentDir -ChildPath $linkName

    Write-Output "Original Working Directory: $originalWorkingDirectory"
    Write-Output "Current Directory: $currentDir"
    Write-Output "Selected Path: $selectedPath"
    Write-Output "Link Path: $linkPath"

	# Check if selected item is a file, and then check if we already have a file with that name and offer to overwrite if we do
	if (Test-Path -Path $selectedPath -PathType Leaf) {
        if (Test-Path -Path $linkPath) {
            $choice = Confirm-Overwrite -linkPath $linkPath
            if ($choice -eq 'Yes') {
                # Remove the existing symlink
                Remove-Item -Path $linkPath -Force
                # Create a new file symbolic link
                New-Item -ItemType SymbolicLink -Path $linkPath -Target $selectedPath -ErrorAction Stop
                Write-Output "Overwritten file symbolic link: $linkPath -> $selectedPath"
            } elseif ($choice -eq 'No') {
                Write-Output "Symbolic link (or file/folder with that name) already existed! Operation canceled by user by not overwriting."
            } else {
                Write-Output "Operation canceled by user."
            }
        } else {
            # Create a new file symbolic link
            New-Item -ItemType SymbolicLink -Path $linkPath -Target $selectedPath -ErrorAction Stop
            Write-Output "Created file symbolic link: $linkPath -> $selectedPath"
        }
    } else {
        Write-Output "The selected path is not a file."
    }
} else {
    Write-Output "No file was selected."
}

# Keep the window open
#Read-Host -Prompt "Press Enter to exit"
Write-Output "`n`nThis window will autoclose soon"
Start-Sleep -Seconds 1
