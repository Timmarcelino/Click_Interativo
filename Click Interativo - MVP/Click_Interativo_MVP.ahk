#Requires AutoHotkey v2.0
#SingleInstance Force

; =========================
; Click Interativo — MVP v1.1.1 (hot fix 1)
; =========================
; - Intervalo por Ponto em **segundos** (UI e INI). Internamente converte para ms.
; - Hot fix: escopo de variáveis globais dentro das funções (Start/Pause/Resume/Stop/SchedulerTick).
; - Restauração opcional de foco & cursor após cada clique por coordenadas.
; - Um SetTimer global (round-robin). Encerramento por data/hora opcional, ou por Esc/erro/janela ausente.
; - Encerrar ciclo NÃO fecha o script.

SendMode "Input"
CoordMode "Mouse", "Client"  ; Click/MouseGetPos ficam relativos à área cliente da janela ativa.

; ---------- Persistência ----------
iniFile := A_ScriptDir "\Click_Interativo.ini"

; ---------- Estado ----------
global gPoints := []            ; {id,title,hwnd,exe,cls,titlePart,coordMode,x,y,useCtrl,classNN,intervalSec,intervalMs,activateBefore,active,nextDue,lastClick}
global gRunning := false, gPaused := false
global gTickMs := 100           ; período do scheduler (não é o intervalo dos pontos)
global gEndAt := ""             ; timestamp "YYYYMMDDHH24MISS" (vazio = sem fim programado)
global gCounters := {clicks: 0, startedAt: 0}

; ---------- GUI principal ----------
g := Gui("", "Click Interativo")
g.SetFont("s10")

; Barra de controlo
btnAdd     := g.Add("Button", "w190", "Configurar Ponto de Interação")
btnTest    := g.Add("Button", "x+6 w130", "Testar Ponto")
btnRemove  := g.Add("Button", "x+6 w130", "Remover Ponto")

; Lista de pontos
lv := g.Add("ListView", "xm w860 h260 Grid", ["ID","Janela/Aplicação","x","y","Modo","ClassNN","Intervalo(s)","Ativar","Ativo"])
lv.ModifyCol(1, 40), lv.ModifyCol(2, 330), lv.ModifyCol(3, 60), lv.ModifyCol(4, 60)
lv.ModifyCol(5, 90), lv.ModifyCol(6, 150), lv.ModifyCol(7, 110), lv.ModifyCol(8, 70), lv.ModifyCol(9, 70)

; Linha de agendamento e flags
cbActivateDefault := g.Add("CheckBox", "xm y+6 Checked", "Ativar janela antes do clique (padrão ao criar Ponto)")
g.Add("Text", "x+12", "Encerrar em:")
dtEnd := g.Add("DateTime", "x+6 w200", "yyyy-MM-dd HH:mm")
cbUseEnd := g.Add("CheckBox", "x+6", "Usar data/hora")

; Restauração de contexto (foco & cursor) para cliques por coordenadas
global cbRestore := g.Add("CheckBox", "xm y+10 Checked", "Restaurar foco & cursor após clique (modo Coordenadas)")

; Controlo do ciclo
btnStart  := g.Add("Button", "xm y+10 w140 Default", "Iniciar Ciclo")
btnPause  := g.Add("Button", "x+6 w120 Disabled", "Pausar")
btnResume := g.Add("Button", "x+6 w120 Disabled", "Retomar")
btnStop   := g.Add("Button", "x+6 w140 Disabled", "Encerrar Ciclo")
lblStatus := g.Add("Text", "xm y+8 w860", "Pronto.")

; Ligações
g.OnEvent("Close", (*) => (gRunning ? StopCycle() : ExitApp()))
btnAdd.OnEvent("Click", AddPointWizard)
btnTest.OnEvent("Click", TestSelectedPoint)
btnRemove.OnEvent("Click", RemoveSelectedPoint)
btnStart.OnEvent("Click", StartCycle)
btnPause.OnEvent("Click", PauseCycle)
btnResume.OnEvent("Click", ResumeCycle)
btnStop.OnEvent("Click", StopCycle)

; Hotkey de encerramento do ciclo (não fecha o script)
Esc:: StopCycle()

g.Show()
LoadPointsFromIni()
RefreshList()

; =========================
;    Implementação
; =========================

; ----- Wizard: Configurar Ponto de Interação -----
AddPointWizard(*) {
    winHwnds := [], winRows := []
    for hwnd in WinGetList() {
        t := ""
        try t := WinGetTitle("ahk_id " hwnd)
        if (t = "")
            continue
        ex := "", cl := ""
        try ex := WinGetProcessName("ahk_id " hwnd)
        try cl := WinGetClass("ahk_id " hwnd)
        winHwnds.Push(hwnd)
        winRows.Push(t "  [" cl "]  (" ex ")")
    }

    w := Gui("+Owner" g.Hwnd, "Configurar Ponto de Interação")
    w.SetFont("s10")

    w.Add("Text",, "Janela/Aplicação:")
    ddl := w.Add("DropDownList", "w620 Choose1", winRows)

    w.Add("Text", "xm y+10", "Coordenadas (Client):")
    edX := w.Add("Edit", "w90 Number", "0")
    edY := w.Add("Edit", "x+8 w90 Number", "0")
    btCap := w.Add("Button", "x+8 w200", "Capturar coordenadas")
    btCap.OnEvent("Click", CapturePoint.Bind(winHwnds, ddl, edX, edY))

    w.Add("Text","xm y+10","Modo de clique:")
    rbCoords := w.Add("Radio", "Group Checked", "Coordenadas (Click)")
    rbCtrl   := w.Add("Radio", "x+12", "ControlClick")
    w.Add("Text","xm","ClassNN (opcional p/ ControlClick):")
    edClass  := w.Add("Edit", "w360", "")

    w.Add("Text","xm y+10","Intervalo por Ponto (segundos):")
    edInterval := w.Add("Edit","w160 Number","1")

    cbAct := w.Add("CheckBox","xm y+10" (cbActivateDefault.Value? " Checked":""), "Ativar janela antes de clicar")
    btnOk := w.Add("Button","xm y+12 w140 Default","Adicionar")
    btnCancel := w.Add("Button","x+8 w140","Cancelar")

    btnCancel.OnEvent("Click", (*) => w.Destroy())
    btnOk.OnEvent("Click", (*) => (
        addPoint(),
        w.Destroy()
    ))

    addPoint() {
        idx := ddl.Value
        if (idx=0)
            return
        hwnd := winHwnds[idx]
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        exe := "", cls := ""
        try exe := WinGetProcessName("ahk_id " hwnd)
        try cls := WinGetClass("ahk_id " hwnd)

        x := Integer(edX.Text), y := Integer(edY.Text)
        if (x = "" || y = "") {
            MsgBox "Defina coordenadas válidas."
            return
        }
        sec := Integer(edInterval.Text)
        if (sec = "" || sec < 1)
            sec := 1
        useCtrl := rbCtrl.Value = 1
        classNN := edClass.Text
        act := cbAct.Value = 1

        p := Map()
        p["id"] := NewPointId()
        p["title"] := title
        p["hwnd"] := hwnd
        p["exe"] := exe
        p["cls"] := cls
        p["titlePart"] := title
        p["coordMode"] := "Client"
        p["x"] := x, p["y"] := y
        p["useCtrl"] := useCtrl, p["classNN"] := classNN
        p["intervalSec"] := sec
        p["intervalMs"] := sec * 1000
        p["activateBefore"] := act
        p["active"] := true
        p["nextDue"] := 0, p["lastClick"] := 0

        gPoints.Push(p)
        SavePointsToIni()
        RefreshList()
    }

    w.Show()
}

; --- Handler: capturar coordenadas ---
CapturePoint(winHwnds, ddl, edX, edY, *) {
    idx := ddl.Value
    if (idx = 0) {
        MsgBox "Escolhe uma janela antes de capturar."
        return
    }
    ActivateWindow("ahk_id " winHwnds[idx])
    ToolTip "Clique e solte no ponto de interação..."
    KeyWait "LButton", "D"
    KeyWait "LButton", "U"
    ToolTip
    MouseGetPos &mx, &my
    edX.Text := mx,  edY.Text := my
}

TestSelectedPoint(*) {
    row := lv.GetNext()
    if (row=0) {
        MsgBox "Escolha um Ponto na lista."
        return
    }
    p := gPoints[row]
    if !EnsureWindow(p) {
        MsgBox "Janela não encontrada para este Ponto."
        return
    }
    DoOneClick(p)
}

RemoveSelectedPoint(*) {
    row := lv.GetNext()
    if (row=0)
        return
    gPoints.RemoveAt(row)
    SavePointsToIni(), RefreshList()
}

; ----- Ciclo de Interação -----
StartCycle(*) {
    global gRunning, gPaused, gCounters, gEndAt, gPoints, btnStart, btnStop, btnPause, btnResume, lblStatus, dtEnd, cbUseEnd
    if (gRunning)
        return
    if (gPoints.Length=0) {
        MsgBox "Adicione pelo menos um Ponto."
        return
    }
    gRunning := true, gPaused := false
    gCounters := {clicks: 0, startedAt: A_TickCount}
    gEndAt := (cbUseEnd.Value=1) ? dtEnd.Value : ""   ; YYYYMMDDHH24MISS vs A_Now. :contentReference[oaicite:2]{index=2}

    now := A_TickCount
    for p in gPoints
        if (p["active"])
            p["nextDue"] := now

    btnStart.Enabled := false, btnStop.Enabled := true
    btnPause.Enabled := true, btnResume.Enabled := false
    lblStatus.Text := "Ciclo em execução…"
    SetTimer SchedulerTick, gTickMs                      ; agenda sem bloquear a GUI. :contentReference[oaicite:3]{index=3}
}

PauseCycle(*) {
    global gRunning, gPaused, btnPause, btnResume, lblStatus
    if (!gRunning || gPaused)
        return
    gPaused := true
    btnPause.Enabled := false, btnResume.Enabled := true
    lblStatus.Text := "Ciclo em pausa."
}

ResumeCycle(*) {
    global gRunning, gPaused, btnPause, btnResume, lblStatus
    if (!gRunning || !gPaused)
        return
    gPaused := false
    btnPause.Enabled := true, btnResume.Enabled := false
    lblStatus.Text := "Ciclo retomado."
}

StopCycle(*) {
    global gRunning, gPaused, gCounters, btnStart, btnStop, btnPause, btnResume, lblStatus
    if (!gRunning) {
        lblStatus.Text := "Parado."
        return
    }
    gRunning := false, gPaused := false
    SetTimer SchedulerTick, 0
    btnStart.Enabled := true, btnStop.Enabled := false
    btnPause.Enabled := false, btnResume.Enabled := false
    durS := Round((A_TickCount - gCounters.startedAt)/1000, 2)
    MsgBox "Ciclo encerrado.`nCliques: " gCounters.clicks "`nDuração: " durS " s"
    lblStatus.Text := "Ciclo encerrado."
}

SchedulerTick(*) {
    global gRunning, gPaused, gEndAt, gPoints, gCounters, lblStatus
    if (!gRunning || gPaused)
        return

    if (gEndAt != "" && A_Now >= gEndAt) {
        StopCycle()
        return
    }

    now := A_TickCount
    for p in gPoints {
        if (!p["active"])
            continue
        if (now < p["nextDue"])
            continue
        if !EnsureWindow(p) {
            lblStatus.Text := "Janela alvo ausente. Encerrando ciclo."
            StopCycle()
            return
        }
        ok := DoOneClick(p)
        if (!ok) {
            lblStatus.Text := "Erro no clique. Encerrando ciclo."
            StopCycle()
            return
        }
        p["lastClick"] := now
        p["nextDue"] := now + p["intervalMs"]   ; agenda próximo vencimento (ms)
        gCounters.clicks += 1
    }
}

; -------- Clique segundo o modo do ponto + restauração de contexto (quando aplicável)
DoOneClick(p) {
    prevCtx := ""
    if (!p["useCtrl"] && cbRestore.Value = 1)
        prevCtx := SaveUserContext()

    success := false
    targetSpec := BuildWinTitle(p)
    targetHwnd := 0
    try {
        if (p["activateBefore"]) {
            if !ActivateWindow(targetSpec)
                throw Error("Falha ao ativar janela alvo.")
        }
        targetHwnd := WinExist(targetSpec)

        if (p["useCtrl"]) {
            SetControlDelay -1                       ; melhora fiabilidade do ControlClick. :contentReference[oaicite:4]{index=4}
            if (p["classNN"]!="")
                ControlClick p["classNN"], targetSpec, , "Left", 1, "NA"
            else
                ControlClick "x" p["x"] " y" p["y"], targetSpec, , "Left", 1, "NA"
        } else {
            Click p["x"], p["y"]                    ; respeita CoordMode "Client"
        }
        success := true
    } catch as e {
        success := false
    }

    if (IsObject(prevCtx))
        RestoreUserContext(prevCtx, targetHwnd)

    return success
}

; ----- Contexto do utilizador: salvar/recuperar janela ativa e cursor (em ecrã)
SaveUserContext() {
    activeHwnd := WinExist("A")                    ; janela ativa (HWND). :contentReference[oaicite:5]{index=5}
    CoordMode "Mouse", "Screen"
    MouseGetPos &sx, &sy
    CoordMode "Mouse", "Client"
    ctx := Map()
    ctx["hwnd"] := activeHwnd
    ctx["x"] := sx, ctx["y"] := sy
    return ctx
}

RestoreUserContext(ctx, targetHwnd := 0) {
    try {
        if (ctx.Has("hwnd") && ctx["hwnd"] && ctx["hwnd"] != targetHwnd && WinExist("ahk_id " ctx["hwnd"])) {
            WinActivate "ahk_id " ctx["hwnd"]
            WinWaitActive "ahk_id " ctx["hwnd"], , 1
        }
        CoordMode "Mouse", "Screen"
        MouseMove ctx["x"], ctx["y"], 0
        CoordMode "Mouse", "Client"
    } catch
        return
}

; ----- Janela alvo / helpers -----
EnsureWindow(p) {
    if (p["hwnd"] && WinExist("ahk_id " p["hwnd"]))
        return true
    spec := BuildWinTitle(p)
    hwnd := WinExist(spec)
    if (hwnd) {
        p["hwnd"] := hwnd
        return true
    }
    return false
}

BuildWinTitle(p) {
    s := ""
    if (p["exe"]!="")
        s .= "ahk_exe " p["exe"] " "
    if (p["cls"]!="")
        s .= "ahk_class " p["cls"] " "
    ; if (p["titlePart"]!="")
    ;     s .= p["titlePart"]
    return RTrim(s)
}

ActivateWindow(winTitleOrHwnd, timeoutMs := 2000) {
    wt := IsInteger(winTitleOrHwnd) ? "ahk_id " winTitleOrHwnd : winTitleOrHwnd
    WinActivate wt
    if WinWaitActive(wt, , timeoutMs/1000)
        return true
    SendEvent "{Alt down}{Alt up}"
    WinActivate wt
    return WinWaitActive(wt, , timeoutMs/1000) != 0
}

; ----- Lista / INI -----
RefreshList() {
    lv.Delete()
    for p in gPoints {
        modo := p["useCtrl"] ? "Control" : "Coords"
        lv.Add(, p["id"], p["title"], p["x"], p["y"], modo, p["classNN"], p["intervalSec"], p["activateBefore"]? "Sim":"Não", p["active"]? "Sim":"Não")
    }
}

NewPointId() {
    max := 0
    for p in gPoints
        if (p["id"]>max)
            max := p["id"]
    return max+1
}

LoadPointsFromIni() {
    if !FileExist(iniFile)
        return
    count := Integer(IniRead(iniFile, "Meta", "Count", "0"))
    Loop count {
        sec := "Point." A_Index
        p := Map()
        p["id"] := Integer(IniRead(iniFile, sec, "id", A_Index))
        p["title"] := IniRead(iniFile, sec, "title", "")
        p["hwnd"] := Integer(IniRead(iniFile, sec, "hwnd", "0"))
        p["exe"] := IniRead(iniFile, sec, "exe", "")
        p["cls"] := IniRead(iniFile, sec, "cls", "")
        p["titlePart"] := IniRead(iniFile, sec, "titlePart", "")
        p["coordMode"] := IniRead(iniFile, sec, "coordMode", "Client")
        p["x"] := Integer(IniRead(iniFile, sec, "x", "0"))
        p["y"] := Integer(IniRead(iniFile, sec, "y", "0"))
        p["useCtrl"] := IniRead(iniFile, sec, "useCtrl", "0") = "1"
        p["classNN"] := IniRead(iniFile, sec, "classNN", "")
        ; compatibilidade: lê 'intervalSec' (novo). Se ausente, tenta 'intervalMs' (antigo).
        intervalSec := Integer(IniRead(iniFile, sec, "intervalSec", ""))
        if (intervalSec = "") {
            oldMs := Integer(IniRead(iniFile, sec, "intervalMs", "1000"))
            intervalSec := Max(1, Round(oldMs/1000))
        }
        p["intervalSec"] := intervalSec
        p["intervalMs"] := intervalSec * 1000
        p["activateBefore"] := IniRead(iniFile, sec, "activateBefore", "1") = "1"
        p["active"] := IniRead(iniFile, sec, "active", "1") = "1"
        p["nextDue"] := 0, p["lastClick"] := 0
        gPoints.Push(p)
    }
}

SavePointsToIni() {
    IniWrite gPoints.Length, iniFile, "Meta", "Count"
    idx := 0
    for p in gPoints {
        idx++
        sec := "Point." idx
        IniWrite p["id"], iniFile, sec, "id"
        IniWrite p["title"], iniFile, sec, "title"
        IniWrite p["hwnd"], iniFile, sec, "hwnd"
        IniWrite p["exe"], iniFile, sec, "exe"
        IniWrite p["cls"], iniFile, sec, "cls"
        IniWrite p["titlePart"], iniFile, sec, "titlePart"
        IniWrite p["coordMode"], iniFile, sec, "coordMode"
        IniWrite p["x"], iniFile, sec, "x"
        IniWrite p["y"], iniFile, sec, "y"
        IniWrite (p["useCtrl"]?1:0), iniFile, sec, "useCtrl"
        IniWrite p["classNN"], iniFile, sec, "classNN"
        IniWrite p["intervalSec"], iniFile, sec, "intervalSec"  ; *** novo padrão em segundos ***
        IniWrite (p["activateBefore"]?1:0), iniFile, sec, "activateBefore"
        IniWrite (p["active"]?1:0), iniFile, sec, "active"
    }
}