#NoEnv
#SingleInstance Force
#Warn
SendMode Input
SetWorkingDir %A_ScriptDir% 

While True {
    if ListenForHotstringStart() {
        ListenForHotstring()
    }
}

; Listens for hotstring trigger chars
ListenForHotstringStart() {
    OutputDebug, Starting to listening for hotstring start
    ih := InputHook("L3 V","{Space}")
    ih.Start()
    ih.Wait()
    Return ih.Input = "!!!"
}

; Listen for actual hotstring
ListenForHotstring() {
    OutputDebug, Listening for hotstring start
    ih := InputHook("V","{Space}")
    ih.Start()
    ih.Wait()
    CheckHotstring(ih.Input)
}

; Check if hotstring is a file in working dir and then do stuff if it is
CheckHotstring(hotstring) {
    OutputDebug % "Received hotstring " . hotstring
    Loop %A_WorkingDir%\*.*
	{
        SplitPath A_LoopFileFullPath,,,, filenameNoExt
        
		if (filenameNoExt = hotstring) {
            OutputDebug % "Found hotstring " . hotstring . " image"
            RemoveWrittenText(hotstring)
            CopyAndPasteFile(A_LoopFileFullPath)
            ListenForHotstringStart()
        }
	}
}

CopyAndPasteFile(fullFileName) {
    InvokeVerb(fullFileName, "Copy")
    Send ^v
}

RemoveWrittenText(hotstring) {
    ; Loop string + space and three exclamations
    Loop % StrLen(hotstring) + 4 {
        SendInput {BackSpace}
    }
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
