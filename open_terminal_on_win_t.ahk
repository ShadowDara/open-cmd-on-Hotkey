; to start the terminal in a folder

#Requires AutoHotkey v2.0
#t:: {
    explorer := ComObject("Shell.Application")
    activeHWND := WinActive("A")

    for window in explorer.Windows {
        try {
            if (window.HWND = activeHWND && InStr(window.FullName, "explorer.exe")) {
                path := window.Document.Folder.Self.Path
                Run Format('cmd.exe /K cd /d "{}"', path)
                return
            }
        }
    }

    MsgBox "No Explorer Window found!."
}
