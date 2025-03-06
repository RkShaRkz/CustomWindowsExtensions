Add-Type -AssemblyName PresentationFramework
# Function to check for elevated privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

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

function Escape-WildcardCharacters {
    param (
        [string]$path
    )
    $escapedPath = $path -replace '([*?\[\]])', '``$1'
    return $escapedPath
}

# Save the initial working directory
$initialWorkingDirectory = [System.IO.Directory]::GetCurrentDirectory()
$escapedInitialWorkingDirectory = Escape-WildcardCharacters -path $initialWorkingDirectory

if (-not (Test-Admin)) {
    Write-Output "Please run this script as an administrator. Let's try rerunning automatically..."
    
    # Re-run the script with elevated privileges, passing the initial working directory as an argument
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$escapedInitialWorkingDirectory`"" -Verb RunAs
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
    }
    else {
        return $null
    }
}

# Function to escape wildcard characters in a path

# Main script
$selectedPath = Show-FilePickerDialog

if ($selectedPath) {
    $currentDir = Get-Location -PSProvider FileSystem
    $linkName = Split-Path -Leaf $selectedPath
    $linkPath = Join-Path -Path $currentDir -ChildPath $linkName

    # Escape wildcard characters in the paths
    $escapedLinkPath = Escape-WildcardCharacters -path $linkPath
    $escapedSelectedPath = Escape-WildcardCharacters -path $selectedPath

    Write-Output "Original Working Directory: $originalWorkingDirectory"
    Write-Output "Current Directory: $currentDir"
    Write-Output "Selected Path: $selectedPath"
    Write-Output "Escaped Selected Path: $escapedSelectedPath"
    Write-Output "Link Path: $linkPath"
    Write-Output "Escaped Link Path: $escapedLinkPath"

    # Check if selected item is a file, and then check if we already have a file with that name and offer to overwrite if we do
    if (Test-Path -Path $escapedSelectedPath -PathType Leaf) {
        if (Test-Path -Path $escapedLinkPath) {
            $choice = Confirm-Overwrite -linkPath $escapedLinkPath
            if ($choice -eq 'Yes') {
                # Remove the existing symlink
                Remove-Item -Path $escapedLinkPath -Force
                # Create a new file symbolic link
                New-Item -ItemType SymbolicLink -Path $linkPath -Target $escapedSelectedPath -ErrorAction Stop
                Write-Output "Overwritten file symbolic link: $escapedLinkPath -> $escapedSelectedPath"
            }
            elseif ($choice -eq 'No') {
                Write-Output "Symbolic link (or file/folder with that name) already existed! Operation canceled by user by not overwriting."
            }
            else {
                Write-Output "Operation canceled by user."
            }
        }
        else {
            # Create a new file symbolic link
            New-Item -ItemType SymbolicLink -Path $linkPath -Target $escapedSelectedPath -ErrorAction Stop
            Write-Output "Created file symbolic link: $escapedLinkPath -> $escapedSelectedPath"
        }
    }
    else {
        Write-Output "The selected path is not a file."
    }
}
else {
    Write-Output "No file was selected."
}

# Keep the window open
#Read-Host -Prompt "Press Enter to exit"
Write-Output "`n`nThis window will autoclose soon"
Start-Sleep -Seconds 1
