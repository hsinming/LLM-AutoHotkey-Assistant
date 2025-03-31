#Requires AutoHotkey v2.0.18+
#Include Dark_MsgBox.ahk ; Enables dark mode MsgBox and InputBox. Remove this if you want light mode MsgBox and InputBox
#Include Dark_Menu.ahk ; Enables dark mode Menu. Remove this if you want light mode Menu
#Include SystemThemeAwareToolTip.ahk ; Enables dark mode tooltips. Remove this if you want light mode tooltips
#Include WebViewToo.ahk ; Allows for use of the WebView2 Framework within AHK to create Web-based GUIs
#Include jsongo.v2.ahk ; For JSON parsing
#Include AutoXYWH.ahk ; Enables auto-resizing of GUI controls. Does not include resizing of Response Window GUI elements, as it is handled by HTML and CSS
#Include ToolTipEx.ahk ; Enables the tooltip to track the mouse cursor smoothly and permit the tooltip to be moved by dragging
DetectHiddenWindows true ; Enables detection of hidden windows for inter-process communication

; ----------------------------------------------------
; Globals
; ----------------------------------------------------

_oldClipboard := ""

; ----------------------------------------------------
; OpenRouter
; ----------------------------------------------------

class OpenRouter {
    static cURLCommand :=
        'cURL.exe -s -X POST https://openrouter.ai/api/v1/chat/completions '
        . '-H "Authorization: Bearer {1}" '
        . '-H "HTTP-Referer: https://github.com/kdalanon/LLM-AutoHotkey-Assistant" '
        . '-H "X-Title: LLM AutoHotkey Assistant" '
        . '-H "Content-Type: application/json" '
        . '-d @"{2}" '
        . '-o "{3}"'

    __New(APIKey) {
        this.APIKey := APIKey
    }

    createJSONRequest(APIModel, systemPrompt, userPrompt) {
        requestObj := {}
        requestObj.model := APIModel
        requestObj.messages := [{
            role: "system",
            content: systemPrompt
        }, {
            role: "user",
            content: userPrompt
        }]
        ; TODO: add other API parameters
        ; requestObj.max_tokens := 4000
        ; requestObj.temperature := 0.8
        return jsongo.Stringify(requestObj)
    }

    extractJSONResponse(var) {
        response := var.Get("choices")[1].Get("message").Get("content")
        
        ; TODO: make a clean () function including the following cleaning process.

        ; remove carriage returns
        response := StrReplace(response, "`r", "")          
        
        ; Remove leading newlines
        while SubStr(response, 1, 1) == '`n' {
            response := SubStr(response, 2)            
        }
        
        ; Remove leading and trailing newlines and spaces
        response := Trim(response)

        ; Recursively remove enclosing double quotes
        while (SubStr(response, 1, 1) == '"' && SubStr(response, -1) == '"') {
            response := SubStr(response, 2, -1)
            response := Trim(response)
        }

        ; Recursively remove enclosing single quotes
        while (SubStr(response, 1, 1) == "'" && SubStr(response, -1) == "'") {
            response := SubStr(response, 2, -1)
            response := Trim(response)
        }

        ; Recursively remove code block backticks
        while (SubStr(response, 1, 1) == "``" && SubStr(response, -1) == "``") {
            response := SubStr(response, 2, -1)
            response := Trim(response)
        }

        ; Change to Windows newline character
        response := StrReplace(response, "`n", "`r`n")

        model := var.Get("model")
        return {
            response: response,
            model: model
        }
    }

    extractErrorResponse(var) {
        error := var.Get("error").Get("message")
        code := var.Get("error").Get("code")
        return {
            error: error,
            code: code,
        }
    }

    appendToChatHistory(role, message, &chatHistoryJSONRequest, chatHistoryJSONRequestFile) {
        obj := jsongo.Parse(chatHistoryJSONRequest)
        obj["messages"].Push({
            role: role,
            content: message
        })
        chatHistoryJSONRequest := jsongo.Stringify(obj)
        FileOpen(chatHistoryJSONRequestFile, "w", "UTF-8-RAW").Write(chatHistoryJSONRequest)
    }

    getMessages(obj) {
        messages := []
        for i in obj["messages"] {
            messages.Push({
                role: i["role"],
                content: i["content"]
            })
        }
        return messages
    }

    removeLastAssistantMessage(&chatHistoryJSONRequest) {
        obj := jsongo.Parse(chatHistoryJSONRequest)
        messagesArray := obj["messages"]
        lastIndex := messagesArray.Length
        if (messagesArray[lastIndex]["role"] = "assistant") {
            messagesArray.RemoveAt(lastIndex)
        }
        chatHistoryJSONRequest := jsongo.Stringify(obj)
    }

    buildcURLCommand(chatHistoryJSONRequestFile, cURLOutputFile) {
        return Format(OpenRouter.cURLCommand, this.APIKey, chatHistoryJSONRequestFile, cURLOutputFile)
    }
}

; ----------------------------------------------------
; Input Window
; ----------------------------------------------------

class InputWindow {
    __New(windowTitle, skipConfirmation := false) {
        this.inputWindowSkipConfirmation := skipConfirmation

        ; Create Input Window
        this.guiObj := Gui("Resize", windowTitle)
        this.guiObj.OnEvent("Close", this.closeButtonAction.Bind(this))
        this.guiObj.OnEvent("Escape", this.closeButtonAction.Bind(this))
        this.guiObj.OnEvent("Size", this.resizeAction.Bind(this))
        this.guiObj.BackColor := "0x212529"
        this.guiObj.SetFont("s14 cWhite", "Cambria")

        ; Add controls
        this.EditControl := this.guiObj.Add("Edit", "x20 y+5 w500 h250 Background0x212529")
        this.SendButton := this.guiObj.Add("Button", "x240 y+10 w80", "Send")

        ; Apply dark mode to title bar
        ; Reference: https://www.autohotkey.com/boards/viewtopic.php?p=422034#p422034
        DllCall("Dwmapi\DwmSetWindowAttribute", "ptr", this.guiObj.hWnd, "int", 20, "int*", true, "int", 4)

        ; Apply dark mode to Send button and Edit control
        for ctrl in [this.SendButton, this.EditControl] {
            DllCall("uxtheme\SetWindowTheme", "ptr", ctrl.hWnd, "str", "DarkMode_Explorer", "ptr", 0)
        }
    }

    showInputWindow(message := "", title := unset, windowID := unset) {
        this.EditControl.Value := message
        if IsSet(title) {
            this.guiObj.Title := title
        }

        this.EditControl.Focus()
        this.guiObj.Show("AutoSize")
        if IsSet(windowID) {
            ControlSend("^{End}", "Edit1", windowID)
        }
    }

    validateInputAndHide(*) {
        if !this.EditControl.Value {
            MsgBox "Please enter a message or close the window.", "No text entered", "IconX"
            return false
        }
        this.guiObj.Hide
        return true
    }

    sendButtonAction(functionToCall) {
        this.SendButton.OnEvent("Click", functionToCall.Bind(this))
    }

    closeButtonAction(*) {
        if this.inputWindowSkipConfirmation || (MsgBox("Close " this.guiObj.Title " window?", this.guiObj.Title, 308) = "Yes") {
            this.EditControl.Value := ""
            this.guiObj.Hide
            return
        }

        return true
    }

    resizeAction(*) {
        AutoXYWH("wh", this.EditControl)
        AutoXYWH("x0.5 y", this.SendButton)
    }

    setSkipConfirmation(value) {
        this.inputWindowSkipConfirmation := value
    }
}

; ----------------------------------------------------
; Custom messages
; ----------------------------------------------------

class CustomMessages {
    static WM_RESPONSE_WINDOW_OPENED := 0x400 + 125
    static WM_RESPONSE_WINDOW_CLOSED := 0x400 + 126
    static WM_SEND_TO_ALL_MODELS := 0x400 + 127
    static WM_RESPONSE_WINDOW_LOADING_START := 0x400 + 123
    static WM_RESPONSE_WINDOW_LOADING_FINISH := 0x400 + 124

    static registerHandlers(origin, handle) {
        switch origin {
            case "mainScript":
                for msg in [this.WM_RESPONSE_WINDOW_OPENED, this.WM_RESPONSE_WINDOW_CLOSED, this.WM_RESPONSE_WINDOW_LOADING_START,
                    this.WM_RESPONSE_WINDOW_LOADING_FINISH]
                    OnMessage(msg, handle)

            case "subScript": OnMessage(this.WM_SEND_TO_ALL_MODELS, handle)
        }
    }

    static notifyResponseWindowState(state, uniqueID, responseWindowhWnd := unset, mainScriptHiddenhWnd := unset) {
        switch state {
            case this.WM_RESPONSE_WINDOW_OPENED, this.WM_RESPONSE_WINDOW_CLOSED:
                PostMessage(state, uniqueID, responseWindowhWnd, , "ahk_id " mainScriptHiddenhWnd)
            case this.WM_SEND_TO_ALL_MODELS:
                PostMessage(state, uniqueID, 0, , "ahk_id " responseWindowhWnd)
            case this.WM_RESPONSE_WINDOW_LOADING_START, this.WM_RESPONSE_WINDOW_LOADING_FINISH:
                PostMessage(state, uniqueID, 0, , "ahk_id " mainScriptHiddenhWnd)
        }
    }
}

; ----------------------------------------------------
; Helper functions
; ----------------------------------------------------

RestoreClipboard() {
    global _oldClipboard
    A_Clipboard := _oldClipboard
    _oldClipboard := ""
}

BackupClipboard() {
    global _oldClipboard
    ; Backup clipboard only if it's not already backed up
    if _oldClipboard == "" {
        _oldClipboard := A_Clipboard
    }
}

GetSelectedTextFromControl() {
    focusedControl := ControlGetFocus("A")  ; Get the ClassNN of the focused control
    if !focusedControl
        return ""  ; No control is focused, return empty string

    hwnd := ControlGetHwnd(focusedControl, "A")  ; Get the HWND of the focused control

    ; Send EM_GETSEL message to get the selection range
    result := DllCall("User32.dll\SendMessageW", "Ptr", hwnd, "UInt", 0xB0, "Ptr", 0, "Ptr", 0, "UInt")

    selStart := result & 0xFFFF  ; Lower 16 bits contain the start index
    selEnd := result >> 16       ; Upper 16 bits contain the end index

    ; Retrieve the full text of the control
    controlText := ControlGetText(hwnd, "A")

    return SubStr(controlText, selStart + 1, selEnd - selStart)
}

GetSelectedText() {    
    ; Initialize text variable
    text := ""

    ; 1. Try copying text using Ctrl+C
    BackupClipboard()
    A_Clipboard := ""
    Send("^c")
    ClipWait(1)
    text := A_Clipboard    
    RestoreClipboard()

    ; 2. If clipboard is empty, try getting selected text from the focused control
    if StrLen(text) < 1
        text := GetSelectedTextFromControl()

    ; 3. If still empty, try getting all text from the focused control
    if StrLen(text) < 1 {
        focusedControl := ControlGetFocus("A")  ; Get focused control's identifier
        if focusedControl
            text := ControlGetText(focusedControl, "A")
    }       
    
    ; De-identify PHI
    ; TODO move De-Identify out of this function.
    ;text := deIdentify(text)

    return text
}

deIdentify(medical_history) {
    ; 1. Names
    medical_history := RegExReplace(medical_history, "[\x{4e00}-\x{9fa5}]{2,4}", "[DE-IDENTIFIED_CHINESE_NAME]") ; Chinese Names
    ;medical_history := RegExReplace(medical_history, "\b[A-Z][a-z]+\s[A-Z][a-z]+\b", "[DE-IDENTIFIED_ENGLISH_NAME]") ; English Names (2 words)
    ;medical_history := RegExReplace(medical_history, "\b[A-Z][a-z]+\s[A-Z][a-z]+\s[A-Z][a-z]+\b", "[DE-IDENTIFIED_ENGLISH_NAME]") ; English Names (3 words)

    ; 2. National Identification Numbers (國民身分證字號)
    medical_history := RegExReplace(medical_history, "\b[A-Z]{1}[12]\d{8}\b", "[DE-IDENTIFIED_NATIONAL_ID]")

    ; 3. Resident Certificate Numbers (居留證號碼)
    medical_history := RegExReplace(medical_history, "\b[A-Z]{2}\d{8}\b", "[DE-IDENTIFIED_RESIDENT_ID]")

    ; 4. Birthdates (出生日期)
    ;medical_history := RegExReplace(medical_history, "\b\d{4}[-/]\d{2}[-/]\d{2}\b", "[DE-IDENTIFIED_BIRTHDATE]") ; YYYY-MM-DD
    ;medical_history := RegExReplace(medical_history, "\b\d{2}[-/]\d{2}[-/]\d{4}\b", "[DE-IDENTIFIED_BIRTHDATE]") ; MM-DD-YYYY
    ;medical_history := RegExReplace(medical_history, "\b\d{2}[-/]\d{2}[-/]\d{2}\b", "[DE-IDENTIFIED_BIRTHDATE]") ; DD-MM-YY (Ambiguous, be careful)
    ;medical_history := RegExReplace(medical_history, "\b(民國|西元)\s*\d{2,3}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\b", "[DE-IDENTIFIED_BIRTHDATE]") ; Chinese format

    ; 5. Phone Numbers (電話號碼)
    medical_history := RegExReplace(medical_history, "(0\d{1,2}-\d{6,8})", "[DE-IDENTIFIED_PHONE]") ; Landlines
    medical_history := RegExReplace(medical_history, "(09\d{2}-\d{3}-\d{3})", "[DE-IDENTIFIED_PHONE]") ; Mobile (with hyphens)
    medical_history := RegExReplace(medical_history, "(09\d{8})", "[DE-IDENTIFIED_PHONE]") ; Mobile (no hyphens)

    ; 6. Email Addresses (電子郵件地址)
    medical_history := RegExReplace(medical_history, "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b", "[DE-IDENTIFIED_EMAIL]")

    ; 7. Medical Record Numbers/Patient IDs (病歷號碼/病人ID) - Example patterns, adjust as needed!
    medical_history := RegExReplace(medical_history, "\b\d{7}\b", "[DE-IDENTIFIED_MEDICAL_ID]") ; 7 digit number
    medical_history := RegExReplace(medical_history, "\b[A-Za-z]\d{6,8}\b", "[DE-IDENTIFIED_MEDICAL_ID]") ; Letter followed by 6-8 digits
    medical_history := RegExReplace(medical_history, "\b[A-Za-z]{2}\d{5,7}\b", "[DE-IDENTIFIED_MEDICAL_ID]") ; Two letters followed by 5-7 digits

    return medical_history
}