; to start the terminal in a folder

#Requires AutoHotkey v2.0
#t:: {
    for window in ComObject("Shell.Application").Windows {
        try {
            if window.HWND = WinActive("A") {
                path := window.Document.Folder.Self.Path
                Run Format('cmd.exe -d "{}"', path)
                return
            }
        }
    }
}
