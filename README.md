# CustomWindowsExtensions
My Custom Windows shell extensions

# SnoozeGuard
Only valid for executable (`*.exe`) files.
When right-clicking an EXE file, this shell extension helps run it with special "priviledges".

The "snooze guard" extension launches the right-clicked executable with one of the following effects:
- prevent display from turning off
- prevent system from going to sleep
- prevent system from going to sleep and display from turning off

Useful for when you just want to run something (that windows can "break" by going to sleep) without
worrying whether that will happen or not.

Whatever the reason, this one can run executables with one of these three effects.

# Command Prompt
Only valid when right-clicking in an empty space in any folder.

Just opens the command prompt in that folder.

# Command Prompt (Admin)
Only valid when right-clicking in an empty space in any folder.

Just opens the command prompt in that folder, as admin, and will ask for UAC elevation

# Paste As
Valid for any file in the clipboard, when right-clicking in empty space in any folder.

Instead of pasting as "OldFile - Copy.<extension>" it will prompt for a name to paste as.
It's just pasting with a customizable name, while keeping the original extension is parametrized and controlled by the checkbox.

In case the checkbox is unchecked, the file will be pasted as-is (that is, as it's written) and omitting the extension in the input
will paste the file with no extension; however changing the extension is possible by just including the new extension in the input itself.

In case the checkbox is checked, if the new file also contains an extension, it will be considered as part of the name, and the
original extension will be put last.

In case a file overwrite needs to happen - either by overwriting some other file or the original itself - you will be presented with a dialog to decide.