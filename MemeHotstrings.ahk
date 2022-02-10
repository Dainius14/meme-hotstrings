#NoEnv
#SingleInstance Force
#Warn
SendMode Input
SetWorkingDir %A_ScriptDir%

global IsListeningForHostring := false
global Hotstring := ""
global AvailableFiles := []
global AvailableFileNamesNoExt := []
global SelectedFiles := []
global SelectedFileNamesNoExt := []
global ih

:*b0:!!!::
    ListenForHotstring()
Return

ListenForHotstring()
{
    if (IsListeningForHostring) {
        Return
    }

    IsListeningForHostring := true
    OutputDebug % "Listening for hotstring..."
    
    ih := InputHook("V T5", "{Esc}")
    ih.OnChar := Func("OnInputHookChar")
    ih.OnKeyDown := Func("OnInputHookKeyDown")
    ih.OnEnd := Func("OnInputHookEnd")
    ih.KeyOpt("{Backspace}", "N")
    ih.KeyOpt("{Tab}", "NSI")
    ih.Start()

    SetAvailableFiles()
    UpdateAfterAction()
}

OnInputHookChar(ih, char)
{
    Hotstring := Hotstring . char
    UpdateAfterAction()
}

OnInputHookKeyDown(ih, vk, sc)
{
    if (vk = 8)
    {
        ; Backspace
        OutputDebug % "Pressed Backspace"

        strLength := StrLen(Hotstring)
        if (strLength > 0) {
            Hotstring := SubStr(Hotstring, 1, strLength - 1)
        }

        UpdateAfterAction()
    }
    else if (vk = 9)
    {
        ; Tab
        OutputDebug % "Pressed Tab"

        if (SelectedFiles.Length() >= 1) {
            DeleteTypedHotstring()
            DeleteTypedTrigger()
            CopyAndPasteFile()
            StopListeningForHotstring()
        }
        ; else if (SelectedFiles.Length() > 1)
        ; {
        ;     DeleteTypedHotstring()
        ;     Hotstring := SelectedFileNamesNoExt[1]
        ;     SendInput % Hotstring
        ;     UpdateAfterAction()
        ; }
    }
}

OnInputHookEnd() {
    StopListeningForHotstring()
}

DeleteTypedHotstring() {
    Loop % StrLen(Hotstring) {
        SendInput {BackSpace}
    }
}

DeleteTypedTrigger() {
    Loop, 3 {
        SendInput {BackSpace}
    }
}

UpdateAfterAction()
{
    OutputDebug % "Current hotstring: " . Hotstring
    SetSelectedFiles()
    ShowTooltip()
}

SetAvailableFiles()
{
    Loop %A_WorkingDir%\*.*
    {
        if (A_LoopFileExt ~= "jpg|png|gif")
        {
            SplitPath A_LoopFileFullPath,,,, fileNameNoExt
            AvailableFiles.Push(A_LoopFileFullPath)
            AvailableFileNamesNoExt.Push(fileNameNoExt)
        }
    }
    OutputDebug % "Set " . AvailableFiles.Length() . " available files"
}

SetSelectedFiles()
{
    SelectedFiles := []
    SelectedFileNamesNoExt := []

    for i, file in AvailableFiles
    {
        fileNameNoExt := AvailableFileNamesNoExt[i]
        if (Hotstring == "" || InStr(fileNameNoExt, Hotstring) == 1)
        {
            SelectedFiles.Push(file)
            SelectedFileNamesNoExt.Push(fileNameNoExt)
        }
    }
    OutputDebug % "Set " . SelectedFiles.Length() . " selected files"
}

ShowTooltip()
{
    tooltipText := ""
    for i, fileName in SelectedFileNamesNoExt
    {
        if (i = 1) {
            tooltipText := fileName . " [TAB]"
        }
        else {
            tooltipText := tooltipText . fileName
        }
        tooltipText := tooltipText . "   "
    }
    tooltipText := SubStr(tooltipText, 1, StrLen(tooltipText) - 3)
    ToolTip % tooltipText, % A_CaretX + 15, % A_CaretY    
}

StopListeningForHotstring()
{
    ih.Stop()
    ToolTip
    IsListeningForHostring := false
    Hotstring :=
    SelectedFiles := []
    SelectedFileNamesNoExt := []
    AvailableFiles := []
    AvailableFileNamesNoExt := []
    OutputDebug % "Stopped listening for hotstring"
}

CopyAndPasteFile() {
    InvokeVerb(SelectedFiles[1], "Copy")
    Send ^v
}

; For copying files
InvokeVerb(path, menu, validate=True) {
    ;by A_Samurai
    ;v 1.0.1 http://sites.google.com/site/ahkref/custom-functions/invokeverb
    objShell := ComObjCreate("Shell.Application")
    if InStr(FileExist(path), "D") || InStr(path, "::{") {
        objFolder := objShell.NameSpace(path) 
        objFolderItem := objFolder.Self
    } else {
        SplitPath, path, name, dir
        objFolder := objShell.NameSpace(dir)
        objFolderItem := objFolder.ParseName(name)
    }
    if validate {
        colVerbs := objFolderItem.Verbs 
        loop % colVerbs.Count {
            verb := colVerbs.Item(A_Index - 1)
            retMenu := verb.name
            StringReplace, retMenu, retMenu, & 
            if (retMenu = menu) {
                verb.DoIt
                Return True
            }
        }
        Return False
    } else
    objFolderItem.InvokeVerbEx(Menu)
}	
