#Requires AutoHotkey v2.0.18+
#SingleInstance

; ----------------------------------------------------
; Global variables
; ----------------------------------------------------

global ActiveModels := Map()

; ----------------------------------------------------
; Include library
; ----------------------------------------------------

#Include lib/Config.ahk
#Include lib/YAML.ahk

; ----------------------------------------------------
; Hotkeys
; ----------------------------------------------------

`:: mainScriptHotkeyActions("showPromptMenu")
~^s:: mainScriptHotkeyActions("saveAndReloadScript")
~^w:: mainScriptHotkeyActions("closeWindows")

#SuspendExempt
CapsLock & `:: mainScriptHotkeyActions("suspendHotkey")

; ----------------------------------------------------
; Functions
; ----------------------------------------------------

; ----------------------------------------------------
; Load multiple YAML files, combine them, and set to prompts.
; ----------------------------------------------------

getPromptArray(yaml_object) {
    prompt_array := []
    for index, yaml_item in yaml_object {
        prompt := Object()
        for key, value in yaml_item {            
            prompt.%key% := value
        }
        prompt_array.Push(prompt)
    }
    return prompt_array
}

LoadYAMLFiles() {
    selectedFiles := []    
    Options := {
        Title: "Select YAML files",
        Filter: "YAML files (*.yaml;*.yml)|*.yaml;*.yml|All files (*.*)|*.*"
    }
    selectedFiles := FileSelect("M3", , Options.Title, Options.Filter)
    
    if !IsSet(selectedFiles) || !selectedFiles.Length
        return ; User cancelled or no files selected

    prompts := [] ; Clear existing prompts
    for filePath in selectedFiles {
        if !FileExist(filePath)
            continue

        try {            
            yaml_object := YAML.parse(FileRead(filePath))
            prompt_array := getPromptArray(yaml_object)
            for prompt in prompt_array {
                prompts.Push(prompt)
            }            
        } catch as e {
            MsgBox("Error loading YAML file: " filePath "`n" e.Message, "Error", "IconX")
        }
    }    
    
    managePromptState("prompts", "set", prompts) ; Update the state with new prompts
}

; ----------------------------------------------------
; Executes main script hotkey actions (menu, suspend, reload, close).
; ----------------------------------------------------

mainScriptHotkeyActions(action) {
    activeModelsCount := ActiveModels.Count

    switch action {
        case "showPromptMenu":
            promptMenu := Menu()
            tagsMap := Map()
            
            ; If active models exist, create "Send message to" submenu with options for each model
            if (activeModelsCount > 0) {
                ; Send message to menu
                sendToMenu := Menu()
                promptMenu.Add("Send message to", sendToMenu)

                for uniqueID, modelData in ActiveModels {
                    sendToMenu.Add(modelData.promptName, sendToPromptGroupHandler.Bind(modelData.promptName))
                }

                ; If there are more than one Response Windows, add "All" menu option
                if (activeModelsCount > 1) {
                    sendToMenu.Add("All", (*) => sendToAllModelsInputWindow.showInputWindow(, , "ahk_id " sendToAllModelsInputWindow.guiObj.hWnd))
                }

                ; Line separator after Activate and Send message to
                promptMenu.Add()
            }

            ; Normal prompts
            prompts := managePromptState("prompts", "get")
            if (IsObject(prompts) && (prompts.Length > 0)) {                
                for index, prompt in prompts {
                    ; Check if prompt has tags
                    hasTags := prompt.HasProp("tags") && prompt.tags && prompt.tags.Length > 0

                    ; If no tags, add directly to menu and continue
                    if !hasTags {
                        promptMenu.Add(prompt.menuText, promptMenuHandler.Bind(index))
                        continue
                    }

                    ; Process tags
                    for tag in prompt.tags {
                        normalizedTag := StrLower(Trim(tag))

                        ; Create tag menu if doesn't exist
                        if !tagsMap.Has(normalizedTag) {
                            tagsMap[normalizedTag] := {menu: Menu(), displayName: tag}
                            promptMenu.Add(tag, tagsMap[normalizedTag].menu)
                        }

                        ; Add prompt to tag menu
                        tagsMap[normalizedTag].menu.Add(prompt.menuText, promptMenuHandler.Bind(index))
                    }
                }
            }

            ; Add menus ("Activate", "Minimize", "Close") that manages Response Windows
            ; after normal prompts if there are active models
            if (activeModelsCount > 0) {
                ; Line separator before managing Response Window menu
                promptMenu.Add()

                ; Define the action types
                actionTypes := ["Activate", "Minimize", "Close"]

                ; Create submenus for each action type
                for _, actionType in actionTypes {
                    ; Convert to lowercase for function names
                    actionKey := StrLower(actionType)

                    actionSubMenu := Menu()
                    promptMenu.Add(actionType, actionSubMenu)

                    ; Add menu items for each active model
                    for uniqueID, modelData in ActiveModels {
                        actionSubMenu.Add(modelData.promptName, managePromptWindows.Bind(actionKey, modelData.promptName
                        ))
                    }

                    ; If there are more than one Response Windows, add "All" menu option
                    if (activeModelsCount > 1) {
                        actionSubMenu.Add("All", managePromptWindows.Bind(actionKey))
                    }
                }
            }

            ; Line separator before Options
            promptMenu.Add()

            ; Options menu
            promptMenu.Add("&Options", optionsMenu := Menu())
            optionsMenu.Add("&1 - Edit prompts", (*) => Run("Notepad " A_ScriptDir "\prompts.yaml"))
            optionsMenu.Add("&2 - Load YAML files", (*) => LoadYAMLFiles())
            optionsMenu.Add("&3 - View available models", (*) => Run("https://openrouter.ai/models"))
            optionsMenu.Add("&4 - View available credits", (*) => Run("https://openrouter.ai/credits"))
            optionsMenu.Add("&5 - View usage activity", (*) => Run("https://openrouter.ai/activity"))
            promptMenu.Show()

        case "suspendHotkey":
            KeyWait "CapsLock", "L"
            SetCapsLockState "Off"
            toggleSuspend(A_IsSuspended)

        case "saveAndReloadScript":
            if !WinActive("prompts.yaml") {
                return
            }

            ; Small delay to ensure file operations are complete
            Sleep 100

            if (activeModelsCount > 0) {
                MsgBox("Script will automatically reload once all Response Windows are closed.", "LLM AutoHotkey Assistant", 64)
                responseWindowState(0, 0, "reloadScript", 0)
            } else {
                Reload()
            }

        case "closeWindows":
            switch WinActive("A") {
                case customPromptInputWindow.guiObj.hWnd: customPromptInputWindow.closeButtonAction()
                case sendToPromptNameInputWindow.guiObj.hWnd: sendToPromptNameInputWindow.closeButtonAction()
                case sendToAllModelsInputWindow.guiObj.hWnd: sendToAllModelsInputWindow.closeButtonAction()
            }
    }
}

; ----------------------------------------------------
; Input Window actions
; ----------------------------------------------------

customPromptSendButtonAction(*) {
    if !customPromptInputWindow.validateInputAndHide() {
        return
    }

    selectedPrompt := managePromptState("selectedPrompt", "get")    
    processInitialRequest(selectedPrompt, customPromptInputWindow.EditControl.Value)
    customPromptInputWindow.EditControl.Value := ""
}

sendToAllModelsSendButtonAction(*) {
    if (ActiveModels.Count = 0) {
        MsgBox "No Response Windows found. Message not sent.", "Send message to all models", "IconX"
        sendToAllModelsInputWindow.guiObj.Hide
        return
    }

    if !sendToAllModelsInputWindow.validateInputAndHide() {
        return
    }

    ; The main script must know each Response Window's JSON file
    ; so it can read it, parse it, append the new
    ; user message, then write it back
    for uniqueID, modelData in ActiveModels {
        JSONStr := FileOpen(modelData.JSONFile, "r", "UTF-8").Read()
        router.appendToChatHistory("user", sendToAllModelsInputWindow.EditControl.Value, &JSONStr, modelData.JSONFile)

        ; Notify the Response Window to re-read the JSON file and call sendRequestToLLM() again
        responseWindowhWnd := modelData.hWnd
        CustomMessages.notifyResponseWindowState(CustomMessages.WM_SEND_TO_ALL_MODELS, uniqueID, responseWindowhWnd)
    }
}

sendToGroupSendButtonAction(*) {
    if (ActiveModels.Count = 0) {
        MsgBox "No Response Windows found. Message not sent.", "Send message to all models", "IconX"
        sendToAllModelsInputWindow.guiObj.Hide
        return
    }

    if !sendToPromptNameInputWindow.validateInputAndHide() {
        return
    }

    if (!targetPromptName := managePromptState("selectedPromptForMessage", "get")) {
        return
    }

    ; Send message only to active models that belong to this prompt
    for uniqueID, modelData in ActiveModels {

        ; Check if this model belongs to the selected prompt
        if (modelData.promptName != targetPromptName) {
            continue
        }

        JSONStr := FileOpen(modelData.JSONFile, "r", "UTF-8").Read()
        router.appendToChatHistory("user", sendToPromptNameInputWindow.EditControl.Value, &JSONStr, modelData.JSONFile)

        ; Notify the Response Window to re-read the JSON file and call sendRequestToLLM() again
        responseWindowhWnd := modelData.hWnd
        CustomMessages.notifyResponseWindowState(CustomMessages.WM_SEND_TO_ALL_MODELS, uniqueID, responseWindowhWnd)
    }

    sendToPromptNameInputWindow.EditControl.Value := ""
}

sendToPromptGroupHandler(promptName, *) {
    promptsList := managePromptState("prompts", "get")

    ; Find the prompt with the matching promptName
    for _, prompt in promptsList {

        ; Check if the prompt has the same name as the one we're looking for
        if (prompt.promptName = promptName) {
            selectedPrompt := prompt
            break
        }
    }

    managePromptState("selectedPromptForMessage", "set", promptName)

    ; Check if the prompt has skipConfirmation property and set accordingly
    sendToPromptNameInputWindow.setSkipConfirmation(selectedPrompt.HasProp("skipConfirmation") ? selectedPrompt.skipConfirmation : false)
    sendToPromptNameInputWindow.showInputWindow(, "Send message to " promptName, "ahk_id " sendToPromptNameInputWindow.guiObj
        .hWnd
    )
}

; Generic function to perform an operation on prompt windows
;
; Parameters:
; - operation (activate, minimize, close): The operation to perform
; - promptName: Optional. If provided, only windows for this prompt will be affected
managePromptWindows(operation, promptName := "", *) {
    ; Create a list of window handles that match our criteria
    hWndsToManage := []

    ; Iterate through all active models
    for uniqueID, modelData in ActiveModels {
        if (promptName = "All" || modelData.promptName = promptName) {
            hWndsToManage.Push(modelData.hWnd)
        }
    }

    ; Perform the requested operation on each window
    for _, hWnd in hWndsToManage {
        switch operation {
            case "activate": WinActivate("ahk_id " hWnd)
            case "minimize": WinMinimize("ahk_id " hWnd)
            case "close": WinClose("ahk_id " hWnd)
        }
    }
}

; ----------------------------------------------------
; Toggle Suspend
; ----------------------------------------------------

toggleSuspend(*) {
    Suspend -1
    if (A_IsSuspended) {
        TraySetIcon("icons\IconOff.ico", , 1)
        A_IconTip := "LLM AutoHotkey Assistant - Suspended)"

        ; Show GUI at the bottom, centered
        scriptSuspendStatus.Show("AutoSize x" (A_ScreenWidth - scriptSuspendStatusWidth) / 2.3 " y990 NA")
    } else {
        TraySetIcon("icons\IconOn.ico")
        A_IconTip := "LLM AutoHotkey Assistant"
        scriptSuspendStatus.Hide()
    }
}

; ----------------------------------------------------
; Prompt menu handler function
; ----------------------------------------------------

promptMenuHandler(index, *) {
    promptsList := managePromptState("prompts", "get")
    selectedPrompt := promptsList[index]
    if (selectedPrompt.HasProp("isCustomPrompt") && selectedPrompt.isCustomPrompt) {
        ; Save the prompt for future reference in customPromptSendButtonAction(*)
        managePromptState("selectedPrompt", "set", selectedPrompt)

        ; Set skipConfirmation property based on the prompt
        customPromptInputWindow.setSkipConfirmation(selectedPrompt.HasProp("skipConfirmation") ? selectedPrompt.skipConfirmation : false)

        customPromptInputWindow.showInputWindow(selectedPrompt.HasProp("customPromptInitialMessage")? selectedPrompt.customPromptInitialMessage : unset, selectedPrompt.promptName, "ahk_id " customPromptInputWindow.guiObj.hWnd)
    } else {
        processInitialRequest(selectedPrompt, selectedPrompt.HasProp("customPromptInitialMessage") ? selectedPrompt.customPromptInitialMessage : unset)
    }
}

; ----------------------------------------------------
; Manage prompt states
; ----------------------------------------------------

managePromptState(component, action, data := {}) {
    static state := {
        prompts: [],
        selectedPrompt: {},
        selectedPromptForMessage: {}
    }

    switch component {
        case "prompts":
            switch action {
                case "get": return state.prompts
                case "set": state.prompts := data
            }

        case "selectedPrompt":
            switch action {
                case "get": return state.selectedPrompt
                case "set": state.selectedPrompt := data
            }

        case "selectedPromptForMessage":
            switch action {
                case "get": return state.selectedPromptForMessage
                case "set": state.selectedPromptForMessage := data
            }
    }
}

; ----------------------------------------------------
; Connect to LLM API and process request
; ----------------------------------------------------

processInitialRequest(prompt, customPromptMessage := unset) {    
    ; Get selected text and active window title
    active_win := WinGetTitle("A")
    selected_text := GetSelectedText()

    if StrLen(selected_text) < 1 and !IsSet(customPromptMessage) {
        manageCursorAndToolTip("Reset")
        MsgBox("The attempt to copy text onto the clipboard failed.", "No text copied", "IconX")
        return
    }

    if IsSet(customPromptMessage) {
        userPrompt := customPromptMessage "`n`n" selected_text    
    } else {
        userPrompt := selected_text
    }

    ; Removes newlines, spaces, and splits by comma
    APIModels := StrSplit(RegExReplace(prompt.APIModels, "\s+", ""), ",")

    ; Automatically disables isAutoPaste if more than one model is present
    isAutoPaste := prompt.HasProp("isAutoPaste") && prompt.isAutoPaste
    isAutoPaste := (APIModels.Length > 1) ? false : isAutoPaste

    for i, fullAPIModelName in APIModels {

        ; Get text before forward slash as providerName
        providerName := SubStr(fullAPIModelName, 1, InStr(fullAPIModelName, "/") - 1)

        ; Get text after forward slash as singleAPIModelName
        singleAPIModelName := SubStr(fullAPIModelName, InStr(fullAPIModelName, "/") + 1)

        uniqueID := A_TickCount

        ; Create the chatHistoryJSONRequest
        chatHistoryJSONRequest := router.createJSONRequest(fullAPIModelName, prompt.systemPrompt, userPrompt)

        ; Generate sanitized filenames for chat history, cURL command, and cURL output files
        chatHistoryJSONRequestFile := A_Temp "\" RegExReplace("chatHistoryJSONRequest_" prompt.promptName "_" singleAPIModelName "_" uniqueID ".json",
            "[\/\\:*?`"<>|]", "")
        cURLCommandFile := A_Temp "\" RegExReplace("cURLCommand_" prompt.promptName "_" singleAPIModelName "_" uniqueID ".txt",
            "[\/\\:*?`"<>|]", "")
        cURLOutputFile := A_Temp "\" RegExReplace("cURLOutput_" prompt.promptName "_" singleAPIModelName "_" uniqueID ".json",
            "[\/\\:*?`"<>|]", "")

        ; Write the JSON request and cURL command to files
        FileOpen(chatHistoryJSONRequestFile, "w", "UTF-8-RAW").Write(chatHistoryJSONRequest)
        cURLCommand := router.buildcURLCommand(chatHistoryJSONRequestFile, cURLOutputFile)
        FileOpen(cURLCommandFile, "w").Write(cURLCommand)

        ; Maintain a reference in the global map
        global ActiveModels        
        ActiveModels[uniqueID] := {
            promptName: prompt.promptName,
            modelName: singleAPIModelName,
            isLoading: false,
            JSONFile: chatHistoryJSONRequestFile,
            ;cURLFile: cURLCommandFile,
            ;outputFile: cURLOutputFile,
            ;provider: router,            
        }

        ; Create an object containing all values for the Response Window
        responseWindowDataObj := {
            chatHistoryJSONRequestFile: chatHistoryJSONRequestFile,
            cURLCommandFile: cURLCommandFile,
            cURLOutputFile: cURLOutputFile,
            providerName: providerName,
            copyAsMarkdown: prompt.HasProp("copyAsMarkdown") && prompt.copyAsMarkdown,
            isAutoPaste: isAutoPaste,
            replaceSelected: prompt.HasProp("replaceSelected") && prompt.replaceSelected,
            responseStart: prompt.HasProp("responseStart") ? prompt.responseStart : "",
            responseEnd: prompt.HasProp("responseEnd") ? Prompt.responseEnd : "",
            skipConfirmation: prompt.HasProp("skipConfirmation") && prompt.skipConfirmation,
            mainScriptHiddenhWnd: A_ScriptHwnd,
            responseWindowTitle: prompt.promptName " [" singleAPIModelName "]",
            singleAPIModelName: singleAPIModelName,
            numberOfAPIModels: APIModels.Length,
            APIModelsIndex: i,
            uniqueID: uniqueID,
            selectedText: selected_text,
            activeWin: active_win
        }

        ; Write the object to a file named responseWindowData and run
        ; Response Window.ahk while passing the location of that file
        ; through dataObjToJSONStrFile as the first argument
        dataObjToJSONStr := jsongo.Stringify(responseWindowDataObj)
        dataObjToJSONStrFile := A_Temp "\" RegExReplace("responseWindowData_" prompt.promptName "_" singleAPIModelName "_" A_TickCount ".json",
            "[\/\\:*?`"<>|]", "")
        FileOpen(dataObjToJSONStrFile, "w", "UTF-8-RAW").Write(dataObjToJSONStr)
        
        ; TODO: delete the following line
        ;ActiveModels[uniqueID].JSONFile := chatHistoryJSONRequestFile
        
        Run("lib\Response Window.ahk " "`"" dataObjToJSONStrFile)
    }
}

; ----------------------------------------------------
; Cursor and Tooltip management
; ----------------------------------------------------

manageCursorAndToolTip(action) {
    switch action {
        case "Update":
            activeCount := 0
            for uniqueID, modelData in ActiveModels {
                if modelData.isLoading {
                    activeCount++
                }
            }

            if (activeCount = 0) {
                ToolTip
                return
            }

            toolTipMessage := "Retrieving response for the following prompt"

            ; Singular and plural forms of the word "model"
            if (activeCount > 1) {
                toolTipMessage .= "s"
            }

            toolTipMessage .= " (Press ESC to cancel):"
            for uniqueID, modelData in ActiveModels {
                if (modelData.isLoading) {
                    toolTipMessage .= "`n- " modelData.promptName " [" modelData.modelName "]"
                }
            }

            ToolTipEX(toolTipMessage, 0)

        case "Loading":
            ; Change default arrow cursor (32512) to "working in background" cursor (32650)
            ; Ensure that other cursors remain unchanged to preserve their functionality
            Cursor := DllCall("LoadCursor", "uint", 0, "uint", 32650)
            DllCall("SetSystemCursor", "Ptr", Cursor, "UInt", 32512)

        case "Reset":
            ToolTip
            DllCall("SystemParametersInfo", "UInt", 0x57, "UInt", 0, "Ptr", 0, "UInt", 0)
    }
}

; ----------------------------------------------------
; Response Window states
; ----------------------------------------------------

responseWindowState(uniqueID, responseWindowhWnd, state, mainScriptHiddenhWnd) {
    global ActiveModels
    static responseWindowLoadingCount := 0
    static reloadScript := false

    switch state {
        case CustomMessages.WM_RESPONSE_WINDOW_OPENED:
            ActiveModels[uniqueID].hWnd := responseWindowhWnd

        case CustomMessages.WM_RESPONSE_WINDOW_CLOSED:
            if ActiveModels.Has(uniqueID) {
                ActiveModels.Delete(uniqueID)
                manageCursorAndToolTip("Update")
            }

            if (ActiveModels.Count = 0) && reloadScript {
                Reload()
            }
        case CustomMessages.WM_RESPONSE_WINDOW_LOADING_START:
            ActiveModels[uniqueID].isLoading := true
            responseWindowLoadingCount++
            if (responseWindowLoadingCount = 1) {
                manageCursorAndToolTip("Loading")
            }

            manageCursorAndToolTip("Update")

        case CustomMessages.WM_RESPONSE_WINDOW_LOADING_FINISH:
            if (responseWindowLoadingCount > 0 && ActiveModels.Has(uniqueID)) {
                responseWindowLoadingCount--
                ActiveModels[uniqueID].isLoading := false
                if (responseWindowLoadingCount = 0) {
                    manageCursorAndToolTip("Reset")
                } else {
                    manageCursorAndToolTip("Update")
                }
            }

        case "reloadScript": reloadScript := true
    }
}

; ----------------------------------------------------
; Auto-execute Section
; ----------------------------------------------------

; ----------------------------------------------------
; API Key Setup: Ask user for OpenRouter API key if not found in settings.ini
; ----------------------------------------------------

if not (FileExist("settings.ini")) {
    api_key := InputBox("Enter your OpenRouter API key", "LLM AutoHotkey Assistant : Setup", "W400 H100").value
    if (api_key == "") {
        MsgBox("To use this script, you need to enter an OpenRouter API key. Please restart the script and try again.")
        ExitApp
    }
    FileCopy("settings.ini.default", "settings.ini")
    IniWrite(api_key, "settings.ini", "settings", "api_key")
}

LoadYAMLFiles()

; ----------------------------------------------------
; Create new instance of OpenRouter class
; ----------------------------------------------------

router := OpenRouter(IniRead("settings.ini", "settings", "api_key"))

; ----------------------------------------------------
; Generate tray menu dynamically
; ----------------------------------------------------

trayMenuItems := [{
    menuText: "&Reload Script",
    function: (*) => Reload()
}, {
    menuText: "E&xit",
    function: (*) => ExitApp()
}]

TraySetIcon("icons\IconOn.ico")
A_TrayMenu.Delete()
for index, item in trayMenuItems {
    A_TrayMenu.Add(item.menuText, item.function)
}
A_IconTip := "LLM AutoHotkey Assistant"

; ----------------------------------------------------
; Create Input Windows
; ----------------------------------------------------

customPromptInputWindow := InputWindow("Custom prompt")
sendToAllModelsInputWindow := InputWindow("Send message to all")
sendToPromptNameInputWindow := InputWindow("Send message to prompt")

; ----------------------------------------------------
; Register sendButtonActions
; ----------------------------------------------------

customPromptInputWindow.sendButtonAction(customPromptSendButtonAction)
sendToAllModelsInputWindow.sendButtonAction(sendToAllModelsSendButtonAction)
sendToPromptNameInputWindow.sendButtonAction(sendToGroupSendButtonAction)

; ----------------------------------------------------
; Initialize Suspend GUI
; ----------------------------------------------------

scriptSuspendStatus := Gui()
scriptSuspendStatus.SetFont("s10", "Cambria")
scriptSuspendStatus.Add("Text", "cBlack Center", "LLM AutoHotkey Assistant Suspended")
scriptSuspendStatus.BackColor := "0xFFDF00"
scriptSuspendStatus.Opt("-Caption +Owner -SysMenu +AlwaysOnTop")
scriptSuspendStatusWidth := ""
scriptSuspendStatus.GetPos(, , &scriptSuspendStatusWidth)

; ----------------------------------------------------
; Custom messages and handlers for detecting
; ----------------------------------------------------

CustomMessages.registerHandlers("mainScript", responseWindowState)
