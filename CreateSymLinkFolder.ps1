$source = @'
// https://stackoverflow.com/a/66823582/21149029
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
/// <summary>
/// Present the Windows Vista-style open file dialog to select a folder. Fall back for older Windows Versions
/// </summary>
#pragma warning disable 0219, 0414, 0162
public class FolderSelectDialog {
    private string _initialDirectory;
    private string _title;
    private string _message;
    private string _fileName = "";
    
    public string InitialDirectory {
        get { return string.IsNullOrEmpty(_initialDirectory) ? Environment.CurrentDirectory : _initialDirectory; }
        set { _initialDirectory = value; }
    }
    public string Title {
        get { return _title ?? "Select a folder"; }
        set { _title = value; }
    }
    public string Message {
        get { return _message ?? _title ?? "Select a folder"; }
        set { _message = value; }
    }
    public string FileName { get { return _fileName; } }

    public FolderSelectDialog(string defaultPath="MyComputer", string title="Select a folder", string message=""){
        InitialDirectory = defaultPath;
        Title = title;
        Message = message;
    }
    
    public bool Show() { return Show(IntPtr.Zero); }

    /// <param name="hWndOwner">Handle of the control or window to be the parent of the file dialog</param>
    /// <returns>true if the user clicks OK</returns>
    public bool Show(IntPtr? hWndOwnerNullable=null) {
        IntPtr hWndOwner = IntPtr.Zero;
        if(hWndOwnerNullable!=null)
            hWndOwner = (IntPtr)hWndOwnerNullable;
        if(Environment.OSVersion.Version.Major >= 6){
            try{
                var resulta = VistaDialog.Show(hWndOwner, InitialDirectory, Title, Message);
                _fileName = resulta.FileName;
                return resulta.Result;
            }
            catch(Exception){
                var resultb = ShowXpDialog(hWndOwner, InitialDirectory, Title, Message);
                _fileName = resultb.FileName;
                return resultb.Result;
            }
        }
        var result = ShowXpDialog(hWndOwner, InitialDirectory, Title, Message);
        _fileName = result.FileName;
        return result.Result;
    }

    private struct ShowDialogResult {
        public bool Result { get; set; }
        public string FileName { get; set; }
    }

    private static ShowDialogResult ShowXpDialog(IntPtr ownerHandle, string initialDirectory, string title, string message) {
        var folderBrowserDialog = new FolderBrowserDialog {
            Description = message,
            SelectedPath = initialDirectory,
            ShowNewFolderButton = true
        };
        var dialogResult = new ShowDialogResult();
        if (folderBrowserDialog.ShowDialog(new WindowWrapper(ownerHandle)) == DialogResult.OK) {
            dialogResult.Result = true;
            dialogResult.FileName = folderBrowserDialog.SelectedPath;
        }
        return dialogResult;
    }

    private static class VistaDialog {
        private const string c_foldersFilter = "Folders|\n";
        
        private const BindingFlags c_flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        private readonly static Assembly s_windowsFormsAssembly = typeof(FileDialog).Assembly;
        private readonly static Type s_iFileDialogType = s_windowsFormsAssembly.GetType("System.Windows.Forms.FileDialogNative+IFileDialog");
        private readonly static MethodInfo s_createVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("CreateVistaDialog", c_flags);
        private readonly static MethodInfo s_onBeforeVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("OnBeforeVistaDialog", c_flags);
        private readonly static MethodInfo s_getOptionsMethodInfo = typeof(FileDialog).GetMethod("GetOptions", c_flags);
        private readonly static MethodInfo s_setOptionsMethodInfo = s_iFileDialogType.GetMethod("SetOptions", c_flags);
        private readonly static uint s_fosPickFoldersBitFlag = (uint) s_windowsFormsAssembly
            .GetType("System.Windows.Forms.FileDialogNative+FOS")
            .GetField("FOS_PICKFOLDERS")
            .GetValue(null);
        private readonly static ConstructorInfo s_vistaDialogEventsConstructorInfo = s_windowsFormsAssembly
            .GetType("System.Windows.Forms.FileDialog+VistaDialogEvents")
            .GetConstructor(c_flags, null, new[] { typeof(FileDialog) }, null);
        private readonly static MethodInfo s_adviseMethodInfo = s_iFileDialogType.GetMethod("Advise");
        private readonly static MethodInfo s_unAdviseMethodInfo = s_iFileDialogType.GetMethod("Unadvise");
        private readonly static MethodInfo s_showMethodInfo = s_iFileDialogType.GetMethod("Show");

        public static ShowDialogResult Show(IntPtr ownerHandle, string initialDirectory, string title, string description) {
            var openFileDialog = new OpenFileDialog {
                AddExtension = false,
                CheckFileExists = false,
                DereferenceLinks = true,
                Filter = c_foldersFilter,
                InitialDirectory = initialDirectory,
                Multiselect = false,
                Title = title
            };

            var iFileDialog = s_createVistaDialogMethodInfo.Invoke(openFileDialog, new object[] { });
            s_onBeforeVistaDialogMethodInfo.Invoke(openFileDialog, new[] { iFileDialog });
            s_setOptionsMethodInfo.Invoke(iFileDialog, new object[] { (uint) s_getOptionsMethodInfo.Invoke(openFileDialog, new object[] { }) | s_fosPickFoldersBitFlag });
            var adviseParametersWithOutputConnectionToken = new[] { s_vistaDialogEventsConstructorInfo.Invoke(new object[] { openFileDialog }), 0U };
            s_adviseMethodInfo.Invoke(iFileDialog, adviseParametersWithOutputConnectionToken);

            try {
                int retVal = (int) s_showMethodInfo.Invoke(iFileDialog, new object[] { ownerHandle });
                return new ShowDialogResult {
                    Result = retVal == 0,
                    FileName = openFileDialog.FileName
                };
            }
            finally {
                s_unAdviseMethodInfo.Invoke(iFileDialog, new[] { adviseParametersWithOutputConnectionToken[1] });
            }
        }
    }

    // Wrap an IWin32Window around an IntPtr
    private class WindowWrapper : IWin32Window {
        private readonly IntPtr _handle;
        public WindowWrapper(IntPtr handle) { _handle = handle; }
        public IntPtr Handle { get { return _handle; } }
    }
    
    public string getPath(){
        if (Show()){
            return FileName;
        }
        return "";
    }
}
'@
# Function to check for elevated privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Confirm-Overwrite {
    param (
        $linkPath
    )
    Add-Type -AssemblyName PresentationFramework

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

# Function to show folder picker dialog
function Show-FolderPickerDialog {

    Add-Type -Language CSharp -TypeDefinition $source -ReferencedAssemblies ("System.Windows.Forms", "System.ComponentModel.Primitives")

    $path = [System.IO.Directory]::GetCurrentDirectory()
    $title = "Select Folder"
    $description = "Select a folder to create a symbolic link"

    $out = [FolderSelectDialog]::new($path, $title, $description).getPath()
    
    if ($out -eq "") {
        return $null
    }
    else {
        return $out
    }
}

# Main script
$selectedPath = Show-FolderPickerDialog

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

    if (Test-Path -Path $escapedSelectedPath -PathType Container) {
        if (Test-Path -Path $escapedLinkPath) {
            $choice = Confirm-Overwrite -linkPath $escapedLinkPath
            if ($choice -eq 'Yes') {
                # Remove the existing symlink
                Remove-Item -Path $escapedLinkPath -Force
                # Create a new directory symbolic link
                New-Item -ItemType SymbolicLink -Path $linkPath -Target $escapedSelectedPath -ErrorAction Stop
                Write-Output "Overwritten directory symbolic link: $escapedLinkPath -> $escapedSelectedPath"
            }
            elseif ($choice -eq 'No') {
                Write-Output "Symbolic link (or file/folder with that name) already existed! Operation canceled by user by not overwriting."
            }
            else {
                Write-Output "Operation canceled by user."
            }
        }
        else {
            # Create a new directory symbolic link
            New-Item -ItemType SymbolicLink -Path $linkPath -Target $escapedSelectedPath -ErrorAction Stop
            Write-Output "Created directory symbolic link: $escapedLinkPath -> $escapedSelectedPath"
        }
    }
    else {
        Write-Output "The selected path is not a folder."
    }
}
else {
    Write-Output "No folder was selected."
}

# Keep the window open
#Read-Host -Prompt "Press Enter to exit"
Write-Output "`n`nThis window will autoclose soon"
Start-Sleep -Seconds 1
