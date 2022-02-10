#NoEnv
#SingleInstance Force
#Warn
SendMode Input
SetWorkingDir %A_ScriptDir%

global TIMEOUT := 5
global APPLY_SUGGESTION_KEY_VK := 0x09
global NEXT_SUGGESTION_KEY_VK := 0xA2

global IsListeningForHostring := false
global Hotstring := ""
global AvailableFiles := []
global AvailableFileNamesNoExt := []
global FilteredFiles := []
global FilteredFileNamesNoExt := []
global SelectedFileIndex := 1
global ih

:*b0:!!!::
    ListenForHotstring()
Return

ListenForHotstring()
{
    if (IsListeningForHostring)
    {
        Return
    }

    IsListeningForHostring := true
    OutputDebug % "Listening for hotstring..."
    
    ih := InputHook("V", "{Esc}")
    ih.Timeout := TIMEOUT
    ih.OnChar := Func("OnInputHookChar")
    ih.OnKeyDown := Func("OnInputHookKeyDown")
    ih.OnEnd := Func("OnInputHookEnd")
    ih.KeyOpt("{Backspace}", "N")
    ih.KeyOpt("{vk" . APPLY_SUGGESTION_KEY_VK . "}{vk" . NEXT_SUGGESTION_KEY_VK . "}", "NSI")
    ih.Start()

    SetAvailableFiles()
    UpdateAfterAction()
}

; *** EVENT LISTENERS
OnInputHookChar(ih, char)
{
    Hotstring := Hotstring . char
    SelectedFileIndex := 1
    UpdateAfterAction()
}

OnInputHookKeyDown(ih, vk, sc)
{
    if (vk = 8)
    {
        OutputDebug % "Pressed Backspace"
        OnBackspacePressed()
    }
    else if (vk = APPLY_SUGGESTION_KEY_VK)
    {
        OutputDebug % "Pressed Apply Suggestion Key"
        OnApplySuggestionPressed()
    }
    else if (vk = NEXT_SUGGESTION_KEY_VK)
    {
        OutputDebug % "Pressed Next Suggestion Key"
        OnNextSuggestionPressed()
    }
    else {
        OutputDebug % "Pressed something: " . vk
    }
}

OnInputHookEnd()
{
    StopListeningForHotstring()
}

OnBackspacePressed()
{
    strLength := StrLen(Hotstring)
    if (strLength > 0)
    {
        Hotstring := SubStr(Hotstring, 1, strLength - 1)
    }

    UpdateAfterAction()
}

OnApplySuggestionPressed()
{
    if (FilteredFiles.Length() = 0)
    {
        Return
    }

    DeleteTypedHotstring()
    DeleteTypedTrigger()
    CopyAndPasteFile()
    StopListeningForHotstring()
}

OnNextSuggestionPressed()
{
    filteredFilesLength := FilteredFiles.Length()

    if (filteredFilesLength = 0)
    {
        Return
    }

    if (SelectedFileIndex + 1 <= filteredFilesLength)
    {
        SelectedFileIndex := SelectedFileIndex + 1
    }
    else
    {
        SelectedFileIndex := 1
    }
    
    UpdateAfterAction()
}
; *** END EVENT LISTENERS

DeleteTypedHotstring()
{
    Loop % StrLen(Hotstring)
    {
        SendInput {BackSpace}
    }
}

DeleteTypedTrigger() {
    Loop, 3
    {
        SendInput {BackSpace}
    }
}

UpdateAfterAction()
{
    OutputDebug % "Current hotstring: " . Hotstring
    SetFilteredFiles()
    ShowTooltip()
    ih.Timeout := TIMEOUT
}

SetAvailableFiles()
{
    Loop %A_WorkingDir%\images\*.*
    {
        SplitPath A_LoopFileFullPath,,,, fileNameNoExt
        AvailableFiles.Push(A_LoopFileFullPath)
        AvailableFileNamesNoExt.Push(fileNameNoExt)
    }
    OutputDebug % "Set " . AvailableFiles.Length() . " available files"
}

SetFilteredFiles()
{
    FilteredFiles := []
    FilteredFileNamesNoExt := []

    for i, file in AvailableFiles
    {
        fileNameNoExt := AvailableFileNamesNoExt[i]
        if (Hotstring == "" || InStr(fileNameNoExt, Hotstring) == 1)
        {
            FilteredFiles.Push(file)
            FilteredFileNamesNoExt.Push(fileNameNoExt)
        }
    }
    OutputDebug % "Set " . FilteredFiles.Length() . " filtered files"
}

ShowTooltip()
{
    tooltipText := ""
    for i, fileName in FilteredFileNamesNoExt
    {
        if (i = SelectedFileIndex)
        {
            StringUpper, upperCaseFileName, fileName
            if (i = 1)
            {
                tooltipText := upperCaseFileName . "  "
            }
            else
            {
                tooltipText := tooltipText . "  " . upperCaseFileName . "  "
            }
        }
        else
        {
            tooltipText := tooltipText . fileName
        }
        tooltipText := tooltipText . "   "
    }
    tooltipText := SubStr(tooltipText, 1, StrLen(tooltipText) - 3)

    if (tooltipText = "")
    {
        tooltipText := "No memes..."
    }

    ToolTip % tooltipText, % A_CaretX + 15, % A_CaretY    
}

StopListeningForHotstring()
{
    ih.Stop()
    ToolTip
    IsListeningForHostring := false
    Hotstring :=
    FilteredFiles := []
    FilteredFileNamesNoExt := []
    SelectedFileIndex := 1

    AvailableFiles := []
    AvailableFileNamesNoExt := []

    ; Reset trigger listener
    Hotstring("Reset")

    OutputDebug % "Stopped listening for hotstring"
}

CopyAndPasteFile()
{
    InvokeVerb(FilteredFiles[SelectedFileIndex], "Copy")
    Send ^v
}

; For copying files
InvokeVerb(path, menu, validate=True)
{
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
