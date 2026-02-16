#NoEnv
#SingleInstance Force
SetBatchLines, -1
FileEncoding, UTF-8

; Admin-панель для управления лицензиями в GitHub (public repo).
; Формат licenses.ini:
;   [Key_<key>]   Status=active|frozen|blocked|deleted   HWIDs=*|hwid1|hwid2
;   [HWID_<hwid>] Status=allowed|blocked

global KP_Ini := A_AppData . "\AutoTyperKeyPanel.ini"
global KP_LastText := ""
global KP_LastSha := ""
global KP_LastError := ""

; Позволяет быстро проверить синтаксис без запуска GUI:
;   AutoHotkey.exe /ErrorStdOut KeyPanel.ahk --syntax-check
try {
    if (A_Args.Length() >= 1 && A_Args[1] = "--syntax-check")
        ExitApp
} catch e {
}

; Defaults (можно менять тут или в GUI)
global KP_Owner := IniGet(KP_Ini, "GitHub", "Owner", "IIo3uTuV4iK")
global KP_Repo := IniGet(KP_Ini, "GitHub", "Repo", "mercurical")
global KP_Branch := IniGet(KP_Ini, "GitHub", "Branch", "main")
global KP_Path := IniGet(KP_Ini, "GitHub", "Path", "licenses.ini")
global KP_Token := IniGet(KP_Ini, "GitHub", "Token", "")

Gui, +Resize +MinSize720x620
Gui, Font, s9, Segoe UI
Gui, Add, GroupBox, x10 y10 w700 h120, GitHub
Gui, Add, Text, x20 y35 w90, Token:
Gui, Add, Edit, vKP_Token x110 y32 w590 Password, %KP_Token%
Gui, Add, Text, x20 y65 w90, Owner:
Gui, Add, Edit, vKP_Owner x110 y62 w160, %KP_Owner%
Gui, Add, Text, x280 y65 w70, Repo:
Gui, Add, Edit, vKP_Repo x350 y62 w170, %KP_Repo%
Gui, Add, Text, x540 y65 w60, Branch:
Gui, Add, Edit, vKP_Branch x600 y62 w100, %KP_Branch%
Gui, Add, Text, x20 y95 w90, Path:
Gui, Add, Edit, vKP_Path x110 y92 w410, %KP_Path%
Gui, Add, Button, x530 y90 w80 h24 gKP_Load, Загрузить
Gui, Add, Button, x615 y90 w85 h24 gKP_Test, Тест
 
Gui, Add, GroupBox, x10 y140 w700 h245, Ключи
Gui, Add, ListView, vKP_KeysLV gKP_KeysLVEvent x20 y165 w680 h185 AltSubmit Grid, Key|Статус|HWIDs|Expires
LV_ModifyCol(1, 180)
LV_ModifyCol(2, 120)
LV_ModifyCol(3, 260)
LV_ModifyCol(4, 90)
Gui, Add, Button, x20 y355 w150 h24 gKP_CreateTempKey, ➕ Временный ключ
Gui, Add, Button, x180 y355 w150 h24 gKP_DeleteSelectedKey, ❌ Удалить выбранный
Gui, Add, Button, x340 y355 w110 h24 gKP_RefreshList, 🔄 Обновить

Gui, Add, GroupBox, x10 y395 w700 h170, Операции (ручной режим)
Gui, Add, Text, x20 y420 w90, Key:
Gui, Add, Edit, vKP_Key x110 y417 w590,
Gui, Add, Text, x20 y450 w90, HWID:
Gui, Add, Edit, vKP_HWID x110 y447 w590,

Gui, Add, Button, x20 y485 w150 h28 gKP_KeyActivate, Key: Активировать
Gui, Add, Button, x180 y485 w150 h28 gKP_KeyFreeze, Key: Заморозить
Gui, Add, Button, x340 y485 w150 h28 gKP_KeyUnfreeze, Key: Разморозить
Gui, Add, Button, x500 y485 w150 h28 gKP_KeyBlock, Key: Заблокировать
Gui, Add, Button, x20 y520 w150 h28 gKP_KeyDelete, Key: Удалить

Gui, Add, Button, x180 y520 w150 h28 gKP_KeyBindHWID, Key: Привязать HWID
Gui, Add, Button, x340 y520 w150 h28 gKP_KeyUnbindHWID, Key: Отвязать HWID
Gui, Add, Button, x500 y520 w150 h28 gKP_KeyAllowAny, Key: HWIDs=*

Gui, Add, Text, vKP_Status x10 y570 w700 h40 cGray, Статус: -

Gui, Show, w720 h620, AutoTyper Key Panel (GitHub)
return

GuiClose:
    ExitApp

KP_Test:
    Gui, Submit, NoHide
    if (KP_Token = "") {
        KP_SetStatus("Token пустой.")
        return
    }
    ok := KP_LoadLicenses(dummyText, dummySha)
    KP_SetStatus(ok ? "OK: доступ к файлу есть." : "Ошибка: " . KP_LastError)
return

KP_Load:
    Gui, Submit, NoHide
    KP_SaveCfg()
    ok := KP_LoadLicenses(text, sha)
    if (!ok) {
        KP_SetStatus("Ошибка: " . KP_LastError)
        return
    }
    KP_LastText := text
    KP_LastSha := sha
    KP_RenderKeys(text)
    KP_SetStatus("Загружено: sha=" . sha . " (" . StrLen(text) . " chars)")
return

KP_RefreshList:
    Gosub, KP_Load
return

KP_DeleteSelectedKey:
    row := LV_GetNext(0, "Focused")
    if (!row)
        row := LV_GetNext()
    if (!row) {
        KP_SetStatus("Не выбран ключ.")
        return
    }
    LV_GetText(k, row, 1)
    if (Trim(k) = "") {
        KP_SetStatus("Не выбран ключ.")
        return
    }
    GuiControl,, KP_Key, %k%
    KP_KeySetStatus("deleted")
    Gosub, KP_Load
return

KP_CreateTempKey:
    Gui, Submit, NoHide
    KP_SaveCfg()
    InputBox, days, Временный ключ, На сколько дней создать ключ? (например 7), , 340, 140
    if (ErrorLevel)
        return
    days := days + 0
    if (days <= 0) {
        KP_SetStatus("Неверное число дней.")
        return
    }
    key := KP_GenKey()
    ts := A_Now
    EnvAdd, ts, %days%, Days
    FormatTime, expires, %ts%, yyyy-MM-dd
    ok := KP_EditAndPush(Func("KP_MutKeySetAll").Bind(key, "active", "*", expires), "Create temp key " . days . "d")
    if (!ok) {
        KP_SetStatus("Ошибка: " . KP_LastError)
        return
    }
    GuiControl,, KP_Key, %key%
    KP_SetStatus("OK: создан временный ключ на " . days . " дней. Expires=" . expires)
    Gosub, KP_Load
return

KP_KeyActivate:
    KP_KeySetStatus("active")
return
KP_KeyFreeze:
    KP_KeySetStatus("frozen")
return
KP_KeyUnfreeze:
    KP_KeySetStatus("active")
return
KP_KeyBlock:
    KP_KeySetStatus("blocked")
return
KP_KeyDelete:
    KP_KeySetStatus("deleted")
return
KP_KeyAllowAny:
    KP_KeySetHWIDs("*")
return

KP_KeyBindHWID:
    Gui, Submit, NoHide
    KP_SaveCfg()
    key := Trim(KP_Key)
    hwid := Trim(KP_HWID)
    if (key = "" || hwid = "") {
        KP_SetStatus("Нужно заполнить Key и HWID.")
        return
    }
    ok := KP_EditAndPush(Func("KP_MutBindHWID").Bind(key, hwid), "Bind HWID")
    KP_SetStatus(ok ? "OK: HWID привязан." : "Ошибка: " . KP_LastError)
return

KP_KeyUnbindHWID:
    Gui, Submit, NoHide
    KP_SaveCfg()
    key := Trim(KP_Key)
    hwid := Trim(KP_HWID)
    if (key = "" || hwid = "") {
        KP_SetStatus("Нужно заполнить Key и HWID.")
        return
    }
    ok := KP_EditAndPush(Func("KP_MutUnbindHWID").Bind(key, hwid), "Unbind HWID")
    KP_SetStatus(ok ? "OK: HWID отвязан." : "Ошибка: " . KP_LastError)
return

KP_HWIDBlock:
    KP_HWIDSetStatus("blocked")
return
KP_HWIDUnblock:
    KP_HWIDSetStatus("allowed")
return

KP_KeysLVEvent:
    if (A_GuiEvent = "DoubleClick") {
        row := A_EventInfo
        if (row) {
            LV_GetText(k, row, 1)
            LV_GetText(stRu, row, 2)
            LV_GetText(hw, row, 3)
            LV_GetText(ex, row, 4)
            KP_OpenEditKeyDialog(k, stRu, hw, ex)
        }
        return
    }
    if (A_GuiEvent = "I") {
        ; selection changed: sync Key field for manual ops
        row := LV_GetNext()
        if (row) {
            LV_GetText(k, row, 1)
            GuiControl,, KP_Key, %k%
        }
    }
return

KP_KeySetStatus(status) {
    global KP_Key, KP_LastError
    Gui, Submit, NoHide
    KP_SaveCfg()
    key := Trim(KP_Key)
    if (key = "") {
        KP_SetStatus("Нужно заполнить Key.")
        return
    }
    ok := KP_EditAndPush(Func("KP_MutKeyStatus").Bind(key, status), "Key status=" . status)
    KP_SetStatus(ok ? "OK: Key status=" . status : "Ошибка: " . KP_LastError)
}

KP_KeySetHWIDs(value) {
    global KP_Key, KP_LastError
    Gui, Submit, NoHide
    KP_SaveCfg()
    key := Trim(KP_Key)
    if (key = "") {
        KP_SetStatus("Нужно заполнить Key.")
        return
    }
    ok := KP_EditAndPush(Func("KP_MutKeyHWIDs").Bind(key, value), "Key HWIDs=" . value)
    KP_SetStatus(ok ? "OK: Key HWIDs=" . value : "Ошибка: " . KP_LastError)
}

KP_HWIDSetStatus(status) {
    global KP_HWID, KP_LastError
    Gui, Submit, NoHide
    KP_SaveCfg()
    hwid := Trim(KP_HWID)
    if (hwid = "") {
        KP_SetStatus("Нужно заполнить HWID.")
        return
    }
    ok := KP_EditAndPush(Func("KP_MutHWIDStatus").Bind(hwid, status), "HWID status=" . status)
    KP_SetStatus(ok ? "OK: HWID status=" . status : "Ошибка: " . KP_LastError)
}

KP_OpenEditKeyDialog(key, stRu, hwids, expires) {
    global KP_Edit_Key, KP_Edit_Status, KP_Edit_HWIDs, KP_Edit_Expires

    KP_Edit_Key := key
    KP_Edit_Status := KP_RuToStatus(stRu)
    KP_Edit_HWIDs := (hwids = "" ? "*" : hwids)
    KP_Edit_Expires := expires

    Gui, KP_Edit:New, +Owner1 +AlwaysOnTop +ToolWindow, Редактировать ключ
    Gui, KP_Edit:Font, s9, Segoe UI
    Gui, KP_Edit:Add, Text, x10 y12 w80, Key:
    Gui, KP_Edit:Add, Edit, vKP_Edit_Key x90 y10 w320 ReadOnly, %KP_Edit_Key%
    Gui, KP_Edit:Add, Text, x10 y42 w80, Статус:
    ; показываем по-русски, сохраняем raw через KP_RuToStatus
    ddl := "Активен|Заморожен|Заблокирован"
    Gui, KP_Edit:Add, DropDownList, vKP_Edit_StatusRu x90 y40 w320, %ddl%
    GuiControl, KP_Edit:ChooseString, KP_Edit_StatusRu, % KP_StatusToRu(KP_Edit_Status)
    Gui, KP_Edit:Add, Text, x10 y72 w80, HWIDs:
    Gui, KP_Edit:Add, Edit, vKP_Edit_HWIDs x90 y70 w320, %KP_Edit_HWIDs%
    Gui, KP_Edit:Add, Text, x10 y102 w80, Expires:
    Gui, KP_Edit:Add, Edit, vKP_Edit_Expires x90 y100 w320, %KP_Edit_Expires%
    Gui, KP_Edit:Add, Text, x90 y124 w320 cGray, Формат: YYYY-MM-DD или пусто (без срока)
    Gui, KP_Edit:Add, Button, x90 y150 w150 h26 gKP_Edit_Save, Сохранить
    Gui, KP_Edit:Add, Button, x260 y150 w150 h26 gKP_Edit_Delete, Удалить
    Gui, KP_Edit:Show, w430 h190
}

KP_Edit_Save:
    global KP_Edit_Key, KP_Edit_HWIDs, KP_Edit_Expires, KP_LastError
    Gui, KP_Edit:Submit, NoHide
    st := KP_RuToStatus(KP_Edit_StatusRu)
    hw := Trim(KP_Edit_HWIDs)
    ex := Trim(KP_Edit_Expires)
    if (hw = "")
        hw := "*"
    if (ex != "" && !RegExMatch(ex, "^(\\d{4}-\\d{2}-\\d{2}|\\d{8}|\\d{14})$")) {
        MsgBox, 48, Ошибка, Expires должен быть YYYY-MM-DD или пустым.
        return
    }
    ok := KP_EditAndPush(Func("KP_MutKeySetAll").Bind(KP_Edit_Key, st, hw, ex), "Edit key " . KP_Edit_Key)
    if (!ok) {
        MsgBox, 48, Ошибка, %KP_LastError%
        return
    }
    Gui, KP_Edit:Destroy
    Gosub, KP_Load
return

KP_Edit_Delete:
    global KP_Edit_Key, KP_LastError
    if (KP_Edit_Key = "")
        return
    ok := KP_EditAndPush(Func("KP_MutKeyStatus").Bind(KP_Edit_Key, "deleted"), "Delete key " . KP_Edit_Key)
    if (!ok) {
        MsgBox, 48, Ошибка, %KP_LastError%
        return
    }
    Gui, KP_Edit:Destroy
    Gosub, KP_Load
return

KP_EditAndPush(mutatorFn, commitMsg) {
    global KP_LastError
    KP_LastError := ""
    if (!KP_LoadLicenses(text, sha))
        return false

    work := A_Temp . "\kp_licenses_work.ini"
    FileDelete, %work%
    FileAppend, %text%, %work%, UTF-8

    ; Apply mutation
    try {
        mutatorFn.Call(work)
    } catch e {
        KP_LastError := "Mutator exception: " . e.Message
        return false
    }

    FileRead, newText, %work%
    newText := KP_NormalizeIniText(newText)
    return KP_PushLicenses(newText, sha, commitMsg)
}

KP_RenderKeys(text) {
    Gui, ListView, KP_KeysLV
    LV_Delete()

    keys := KP_ParseKeysFromIniText(text)
    for _, obj in keys {
        stRu := KP_StatusToRu(obj.Status)
        hw := (obj.HWIDs = "" ? "*" : obj.HWIDs)
        ex := obj.Expires
        LV_Add("", obj.Key, stRu, hw, ex)
    }
    LV_ModifyCol()
}

KP_ParseKeysFromIniText(text) {
    arr := []
    cur := ""
    obj := ""

    ; Normalize newlines
    t := StrReplace(text, "`r`n", "`n")
    t := StrReplace(t, "`r", "`n")
    lines := StrSplit(t, "`n")

    for _, line in lines {
        line := Trim(line, " `t")
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue

        if (RegExMatch(line, "i)^\\[\\s*key_(.+?)\\s*\\]$", m)) {
            ; push prev
            if (IsObject(obj))
                arr.Push(obj)
            cur := m1
            obj := {Key: cur, Status: "", HWIDs: "*", Expires: ""}
            continue
        }
        if (RegExMatch(line, "i)^\\[.+\\]$")) {
            if (IsObject(obj))
                arr.Push(obj)
            obj := ""
            continue
        }
        if (!IsObject(obj))
            continue

        if (RegExMatch(line, "i)^status\\s*=\\s*(.*)$", m2)) {
            obj.Status := StrLower(Trim(m21))
            continue
        }
        if (RegExMatch(line, "i)^hwids\\s*=\\s*(.*)$", m3)) {
            obj.HWIDs := Trim(m31)
            continue
        }
        if (RegExMatch(line, "i)^expires\\s*=\\s*(.*)$", m4)) {
            obj.Expires := Trim(m41)
            continue
        }
    }
    if (IsObject(obj))
        arr.Push(obj)
    return arr
}

KP_StatusToRu(st) {
    st := StrLower(Trim(st))
    if (st = "active")
        return "Активен"
    if (st = "frozen")
        return "Заморожен"
    if (st = "blocked")
        return "Заблокирован"
    if (st = "deleted")
        return "Удалён"
    if (st = "")
        return "-"
    return st
}

KP_RuToStatus(stRu) {
    s := StrLower(Trim(stRu))
    if (s = "активен")
        return "active"
    if (s = "заморожен")
        return "frozen"
    if (s = "заблокирован")
        return "blocked"
    if (s = "удалён" || s = "удален")
        return "deleted"
    ; already raw
    return s
}

KP_LoadLicenses(ByRef text, ByRef sha) {
    global KP_Owner, KP_Repo, KP_Branch, KP_Path, KP_Token, KP_LastError
    KP_LastError := ""
    if (KP_Token = "") {
        KP_LastError := "Token пустой."
        return false
    }
    url := "https://api.github.com/repos/" . KP_Owner . "/" . KP_Repo . "/contents/" . KP_Path . "?ref=" . KP_Branch
    ok := KP_Http("GET", url, KP_Token, "", st, body)
    if (!ok) {
        KP_LastError := "HTTP error"
        return false
    }
    if (st < 200 || st >= 300) {
        KP_LastError := "GitHub HTTP " . st . ": " . body
        return false
    }
    sha := KP_JsonGet(body, "sha")
    b64 := KP_JsonGet(body, "content")
    if (sha = "" || b64 = "") {
        KP_LastError := "sha/content не найдено"
        return false
    }
    b64 := KP_JsonUnescape(b64)
    b64 := StrReplace(b64, "`n")
    b64 := StrReplace(b64, "`r")
    b64 := StrReplace(b64, "\n")
    b64 := StrReplace(b64, "\r")
    b64 := RegExReplace(b64, "\\s+")
    bytes := KP_B64ToBytes(b64)
    text := KP_BytesToTextUtf8(bytes)
    return true
}

KP_PushLicenses(text, sha, commitMsg) {
    global KP_Owner, KP_Repo, KP_Branch, KP_Path, KP_Token, KP_LastError
    KP_LastError := ""

    bytes := KP_TextToBytesUtf8(text)
    b64 := KP_BytesToB64(bytes)
    msg := (commitMsg != "") ? commitMsg : "Update licenses"

    url := "https://api.github.com/repos/" . KP_Owner . "/" . KP_Repo . "/contents/" . KP_Path
    json := "{"
        . """message"":""" . KP_JsonEscape(msg) . ""","
        . """content"":""" . b64 . ""","
        . """sha"":""" . KP_JsonEscape(sha) . ""","
        . """branch"":""" . KP_JsonEscape(KP_Branch) . """"
        . "}"

    ok := KP_Http("PUT", url, KP_Token, json, st, body)
    if (!ok) {
        KP_LastError := "HTTP error"
        return false
    }
    if (st < 200 || st >= 300) {
        KP_LastError := "GitHub HTTP " . st . ": " . body
        return false
    }
    return true
}

; === Mutators ===
KP_MutKeyStatus(key, status, filePath) {
    sec := "Key_" . KP_SanitizeIniSection(key)
    status := StrLower(Trim(status))
    if (status = "deleted") {
        IniDelete, %filePath%, %sec%
        return
    }
    IniWrite, %status%, %filePath%, %sec%, Status
    ; Ensure HWIDs exists
    IniRead, hw, %filePath%, %sec%, HWIDs, *
    if (Trim(hw) = "")
        IniWrite, *, %filePath%, %sec%, HWIDs
}

KP_MutKeySetAll(key, status, hwids, expires, filePath) {
    sec := "Key_" . KP_SanitizeIniSection(key)
    st := StrLower(Trim(status))
    if (st = "" || st = "deleted") {
        IniDelete, %filePath%, %sec%
        return
    }
    if (st != "active" && st != "frozen" && st != "blocked")
        st := "active"
    IniWrite, %st%, %filePath%, %sec%, Status

    hw := Trim(hwids)
    if (hw = "")
        hw := "*"
    IniWrite, %hw%, %filePath%, %sec%, HWIDs

    ex := Trim(expires)
    if (ex = "") {
        IniDelete, %filePath%, %sec%, Expires
    } else {
        IniWrite, %ex%, %filePath%, %sec%, Expires
    }
}

KP_MutKeyHWIDs(key, value, filePath) {
    sec := "Key_" . KP_SanitizeIniSection(key)
    v := Trim(value)
    if (v = "")
        v := "*"
    IniWrite, %v%, %filePath%, %sec%, HWIDs
    IniRead, st, %filePath%, %sec%, Status, active
    if (Trim(st) = "")
        IniWrite, active, %filePath%, %sec%, Status
}

KP_MutBindHWID(key, hwid, filePath) {
    sec := "Key_" . KP_SanitizeIniSection(key)
    h := Trim(hwid)
    if (h = "")
        return
    IniRead, hw, %filePath%, %sec%, HWIDs, *
    hw := Trim(hw)
    if (hw = "*" || hw = "") {
        IniWrite, %h%, %filePath%, %sec%, HWIDs
        return
    }
    if (KP_ListHas(hw, h))
        return
    hw := hw . "|" . h
    IniWrite, %hw%, %filePath%, %sec%, HWIDs
}

KP_MutUnbindHWID(key, hwid, filePath) {
    sec := "Key_" . KP_SanitizeIniSection(key)
    h := Trim(hwid)
    if (h = "")
        return
    IniRead, hw, %filePath%, %sec%, HWIDs, *
    hw := Trim(hw)
    if (hw = "*" || hw = "")
        return
    out := ""
    Loop, Parse, hw, |
    {
        one := Trim(A_LoopField)
        if (one = "" || one = h)
            continue
        out .= (out = "" ? "" : "|") . one
    }
    if (out = "")
        out := "*"
    IniWrite, %out%, %filePath%, %sec%, HWIDs
}

KP_MutHWIDStatus(hwid, status, filePath) {
    sec := "HWID_" . KP_SanitizeIniSection(hwid)
    st := StrLower(Trim(status))
    if (st != "blocked")
        st := "allowed"
    IniWrite, %st%, %filePath%, %sec%, Status
}

; === Helpers ===
KP_SaveCfg() {
    global KP_Ini
    Gui, Submit, NoHide
    IniWrite, %KP_Owner%, %KP_Ini%, GitHub, Owner
    IniWrite, %KP_Repo%, %KP_Ini%, GitHub, Repo
    IniWrite, %KP_Branch%, %KP_Ini%, GitHub, Branch
    IniWrite, %KP_Path%, %KP_Ini%, GitHub, Path
    IniWrite, %KP_Token%, %KP_Ini%, GitHub, Token
}

KP_SetStatus(s) {
    GuiControl,, KP_Status, % "Статус: " . s
}

IniGet(file, section, key, def := "") {
    IniRead, v, %file%, %section%, %key%, %def%
    return v
}

KP_SanitizeIniSection(s) {
    s := Trim(s)
    if (s = "")
        return "EMPTY"
    s := RegExReplace(s, "[^0-9A-Za-z_-]", "_")
    s := RegExReplace(s, "_{2,}", "_")
    s := Trim(s, "_")
    if (StrLen(s) > 120)
        s := SubStr(s, 1, 120)
    return s
}

StrLower(s) {
    StringLower, out, s
    return out
}

KP_ListHas(list, value) {
    v := Trim(value)
    if (v = "")
        return false
    Loop, Parse, list, |
    {
        if (Trim(A_LoopField) = v)
            return true
    }
    return false
}

KP_NormalizeIniText(text) {
    ; Ensure CRLF for GitHub diffs + keep UTF-8.
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    text := StrReplace(text, "`n", "`r`n")
    return text
}

KP_Http(method, url, token, body, ByRef status, ByRef respText) {
    respText := ""
    status := 0
    try {
        http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        http.Open(method, url, false)
        http.SetRequestHeader("User-Agent", "AutoTyper-KeyPanel")
        http.SetRequestHeader("Accept", "application/vnd.github+json")
        if (token != "")
            http.SetRequestHeader("Authorization", "Bearer " . token)
        if (method = "PUT" || method = "POST" || method = "PATCH") {
            http.SetRequestHeader("Content-Type", "application/json")
        }
        http.Send(body)
        status := http.Status + 0
        respText := KP_HttpResponseTextUtf8(http)
        return true
    } catch e {
        return false
    }
}

KP_HttpResponseTextUtf8(http) {
    try bytes := http.ResponseBody
    catch e
        return http.ResponseText
    stream := ComObjCreate("ADODB.Stream")
    stream.Type := 1
    stream.Open()
    stream.Write(bytes)
    stream.Position := 0
    stream.Type := 2
    stream.Charset := "utf-8"
    text := stream.ReadText()
    stream.Close()
    return text
}

KP_JsonGet(json, key) {
    ; Very small JSON getter for "key":"value" (top-level).
    if (RegExMatch(json, "s)""" . key . """\\s*:\\s*""((?:\\\\.|[^""])*)""", m))
        return m1
    return ""
}

KP_JsonEscape(s) {
    bs := Chr(92), dq := Chr(34)
    s := StrReplace(s, bs, bs . bs)
    s := StrReplace(s, dq, bs . dq)
    s := StrReplace(s, "`r", "\\r")
    s := StrReplace(s, "`n", "\\n")
    s := StrReplace(s, "`t", "\\t")
    return s
}

KP_JsonUnescape(s) {
    ; Minimal unescape: \\n, \\r, \\t, \\\\, \\"
    bs := Chr(92), dq := Chr(34)
    s := StrReplace(s, "\\n", "`n")
    s := StrReplace(s, "\\r", "`r")
    s := StrReplace(s, "\\t", "`t")
    s := StrReplace(s, bs . dq, dq)
    s := StrReplace(s, bs . bs, bs)
    return s
}

KP_B64ToBytes(b64) {
    doc := ComObjCreate("MSXML2.DOMDocument.6.0")
    node := doc.createElement("b64")
    node.dataType := "bin.base64"
    node.text := b64
    return node.nodeTypedValue
}

KP_BytesToB64(bytes) {
    doc := ComObjCreate("MSXML2.DOMDocument.6.0")
    node := doc.createElement("b64")
    node.dataType := "bin.base64"
    node.nodeTypedValue := bytes
    b64 := node.text
    return RegExReplace(b64, "\\s+")
}

KP_BytesToTextUtf8(bytes) {
    stream := ComObjCreate("ADODB.Stream")
    stream.Type := 1
    stream.Open()
    stream.Write(bytes)
    stream.Position := 0
    stream.Type := 2
    stream.Charset := "utf-8"
    text := stream.ReadText()
    stream.Close()
    return text
}

KP_TextToBytesUtf8(text) {
    stream := ComObjCreate("ADODB.Stream")
    stream.Type := 2
    stream.Charset := "utf-8"
    stream.Open()
    stream.WriteText(text)
    stream.Position := 0
    stream.Type := 1
    bytes := stream.Read()
    stream.Close()
    return bytes
}

KP_GenKey() {
    ; Простая генерация ключа (не криптостойко; repo public).
    chars := "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    out := ""
    Loop, 20 {
        Random, r, 1, % StrLen(chars)
        out .= SubStr(chars, r, 1)
        if (Mod(A_Index, 5) = 0 && A_Index < 20)
            out .= "-"
    }
    return out
}
