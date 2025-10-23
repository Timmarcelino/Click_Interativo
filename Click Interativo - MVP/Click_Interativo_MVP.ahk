#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================
; Click Interativo — MVP v2.0.2
; ============================================
; v2.0:  Contador regressivo (“Falta(s)”), drift px (↑→↓←), jitter ±s alternado,
;        feedback visual no ponto clicado.
; v2.0.1: Corrige 'Unexpected "{"' no EnsureWindow() (bloco multi-linha).
; v2.0.2: Corrige atualização da célula do ListView (usar lv.Modify em vez de SetText).
; Mantém: 1 SetTimer global; fim por data/hora opcional; restauração de foco & cursor;
;         persistência INI; intervalos em segundos.

SendMode "Input"
CoordMode "Mouse", "Client"

; ---------- Persistência ----------
iniFile := A_ScriptDir "\Click_Interativo.ini"

; ---------- Estado ----------
global gPoints := []  ; {id,title,hwnd,exe,cls,titlePart,coordMode,x,y,useCtrl,classNN,intervalSec,intervalMs,activateBefore,active,nextDue,lastClick,driftStep,jitterPhase}
global gRunning := false, gPaused := false
global gTickMs := 100
global gEndAt := ""             ; YYYYMMDDHH24MISS
global gCounters := {clicks: 0, startedAt: 0}
global gLastCountdownRefresh := 0
global gJitterSec := 1

; ---------- GUI principal ----------
g := Gui("", "Click Interativo")
g.SetFont("s10")

btnAdd     := g.Add("Button", "w190", "Configurar Ponto de Interação")
btnTest    := g.Add("Button", "x+6 w130", "Testar Ponto")
btnRemove  := g.Add("Button", "x+6 w130", "Remover Ponto")

; 1:ID 2:Janela 3:x 4:y 5:Modo 6:ClassNN 7:Intervalo(s) 8:Falta(s) 9:Ativar 10:Ativo
lv := g.Add("ListView", "xm w940 h260 Grid", ["ID","Janela/Aplicação","x","y","Modo","ClassNN","Intervalo(s)","Falta(s)","Ativar","Ativo"])
lv.ModifyCol(1, 40), lv.ModifyCol(2, 330), lv.ModifyCol(3, 60), lv.ModifyCol(4, 60)
lv.ModifyCol(5, 90), lv.ModifyCol(6, 150), lv.ModifyCol(7, 100), lv.ModifyCol(8, 90), lv.ModifyCol(9, 70), lv.ModifyCol(10, 70)

cbActivateDefault := g.Add("CheckBox", "xm y+6 Checked", "Ativar janela antes do clique (padrão ao criar Ponto)")
g.Add("Text", "x+12", "Encerrar em:")
dtEnd := g.Add("DateTime", "x+6 w200", "yyyy-MM-dd HH:mm")
cbUseEnd := g.Add("CheckBox", "x+6", "Usar data/hora")

global cbRestore := g.Add("CheckBox", "xm y+10 Checked", "Restaurar foco & cursor após clique (modo Coordenadas)")
g.Add("Text", "x+16", "Jitter (±s):")
edJitter := g.Add("Edit", "x+6 w60 Number", "1")

btnStart  := g.Add("Button", "xm y+10 w150 Default", "Iniciar Ciclo")
btnPause  := g.Add("Button", "x+6 w120 Disabled", "Pausar")
btnResume := g.Add("Button", "x+6 w120 Disabled", "Retomar")
btnStop   := g.Add("Button", "x+6 w150 Disabled", "Encerrar Ciclo")
lblStatus := g.Add("Text", "xm y+8 w940", "Pronto.")

g.OnEvent("Close", (*) => (gRunning ? StopCycle() : ExitApp()))
btnAdd.OnEvent("Click", AddPointWizard)
btnTest.OnEvent("Click", TestSelectedPoint)
btnRemove.OnEvent("Click", RemoveSelectedPoint)
btnStart.OnEvent("Click", StartCycle)
btnPause.OnEvent("Click", PauseCycle)
btnResume.OnEvent("Click", ResumeCycle)
btnStop.OnEvent("Click", StopCycle)

Esc:: StopCycle()

g.Show()
LoadPointsFromIni()
RefreshList()

; ============================================
;              Implementação
; ============================================

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
        winHwnds.Push(hwnd), winRows.Push(t "  [" cl "]  (" ex ")")
    }

    w := Gui("+Owner" g.Hwnd, "Configurar Ponto de Interação")
    w.SetFont("s10")

    w.Add("Text",, "Janela/Aplicação:")
    ddl := w.Add("DropDownList", "w700 Choose1", winRows)

    w.Add("Text", "xm y+10", "Coordenadas (Client):")
    edX := w.Add("Edit", "w90 Number", "0")
    edY := w.Add("Edit", "x+8 w90 Number", "0")
    btCap := w.Add("Button", "x+8 w200", "Capturar coordenadas")
    btCap.OnEvent("Click", CapturePoint.Bind(winHwnds, ddl, edX, edY))

    w.Add("Text","xm y+10","Modo de clique:")
    rbCoords := w.Add("Radio", "Group Checked", "Coordenadas (Click)")
    rbCtrl   := w.Add("Radio", "x+12", "ControlClick")
    w.Add("Text","xm","ClassNN (opcional p/ ControlClick):")
    edClass  := w.Add("Edit", "w380", "")

    w.Add("Text","xm y+10","Intervalo por Ponto (segundos):")
    edInterval := w.Add("Edit","w160 Number","1")

    cbAct := w.Add("CheckBox","xm y+10" (cbActivateDefault.Value? " Checked":""), "Ativar janela antes de clicar")
    btnOk := w.Add("Button","xm y+12 w160 Default","Adicionar")
    btnCancel := w.Add("Button","x+8 w140","Cancelar")

    btnCancel.OnEvent("Click", (*) => w.Destroy())
    btnOk.OnEvent("Click", (*) => (addPoint(), w.Destroy()))

    addPoint() {
        idx := ddl.Value
        if (idx=0)
            return
        hwnd := winHwnds[idx]
        title := "", exe := "", cls := ""
        try title := WinGetTitle("ahk_id " hwnd)
        try exe := WinGetProcessName("ahk_id " hwnd)
        try cls := WinGetClass("ahk_id " hwnd)

        x := Integer(edX.Text), y := Integer(edY.Text)
        if (x = "" || y = "") {
            MsgBox "Defina coordenadas válidas."
            return
        }
        secText := Trim(edInterval.Text)
        sec := (IsNumber(secText) ? Max(1, Integer(secText)) : 1)

        p := Map()
        p["id"] := NewPointId()
        p["title"] := title
        p["hwnd"] := hwnd
        p["exe"] := exe
        p["cls"] := cls
        p["titlePart"] := title
        p["coordMode"] := "Client"
        p["x"] := x, p["y"] := y
        p["useCtrl"] := (rbCtrl.Value = 1)
        p["classNN"] := edClass.Text
        p["intervalSec"] := sec
        p["intervalMs"] := sec * 1000
        p["activateBefore"] := (cbAct.Value = 1)
        p["active"] := true
        p["nextDue"] := 0, p["lastClick"] := 0
        p["driftStep"] := 0
        p["jitterPhase"] := 0

        gPoints.Push(p), SavePointsToIni(), RefreshList()
    }

    w.Show()
}

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
    global gRunning, gPaused, gCounters, gEndAt, gPoints, btnStart, btnStop, btnPause, btnResume, lblStatus, dtEnd, cbUseEnd, edJitter, gJitterSec
    if (gRunning)
        return
    if (gPoints.Length=0) {
        MsgBox "Adicione pelo menos um Ponto."
        return
    }
    jt := Trim(edJitter.Text)
    gJitterSec := (IsNumber(jt) ? Abs(Integer(jt)) : 1)

    gRunning := true, gPaused := false
    gCounters := {clicks: 0, startedAt: A_TickCount}
    gEndAt := (cbUseEnd.Value=1) ? dtEnd.Value : ""

    now := A_TickCount
    for p in gPoints {
        if (p["active"]) {
            p["nextDue"] := now
            p["driftStep"] := 0
            p["jitterPhase"] := 0
        }
    }

    btnStart.Enabled := false, btnStop.Enabled := true
    btnPause.Enabled := true, btnResume.Enabled := false
    lblStatus.Text := "Ciclo em execução…"
    SetTimer SchedulerTick, gTickMs
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
    global gRunning, gPaused, gEndAt, gPoints, gCounters, lblStatus, gLastCountdownRefresh, lv
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

        ; próximo vencimento com jitter alternado ±
        nextMs := p["intervalMs"]
        if (gJitterSec > 0) {
            jitter := gJitterSec * 1000
            nextMs += (p["jitterPhase"] ? jitter : -jitter)
            p["jitterPhase"] := !p["jitterPhase"]
        }
        nextMs := Max(100, nextMs)
        p["nextDue"] := now + nextMs
        gCounters.clicks += 1
        lblStatus.Text := "Clique em: " p["title"] "  (" gCounters.clicks " total)"
    }

    ; atualizar “Falta(s)” ~4x/seg
    if (now - gLastCountdownRefresh >= 250) {
        gLastCountdownRefresh := now
        idx := 0
        for p in gPoints {
            idx++
            if (!p["active"]) {
                lv.Modify(idx, "Col8", "")
                continue
            }
            rem := Max(0, p["nextDue"] - now) / 1000.0
            lv.Modify(idx, "Col8", Format("{:.1f}", rem))  ; atualizar célula (ListView.Modify + "ColN"). :contentReference[oaicite:2]{index=2}
        }
    }
}

; -------- Clique + drift + feedback + restauração (quando aplicável)
DoOneClick(p) {
    prevCtx := ""
    if (!p["useCtrl"] && cbRestore.Value = 1)
        prevCtx := SaveUserContext()

    success := false
    targetSpec := BuildWinTitle(p)
    targetHwnd := 0

    drift := [[0,-1],[1,0],[0,1],[-1,0]]
    dx := 0, dy := 0
    if (!p["useCtrl"]) {
        step := Mod(p["driftStep"], 4)
        dx := drift[step+1][1], dy := drift[step+1][2]
    }

    try {
        if (p["activateBefore"]) {
            if !ActivateWindow(targetSpec)
                throw Error("Falha ao ativar janela alvo.")
        }
        targetHwnd := WinExist(targetSpec)

        if (p["useCtrl"]) {
            SetControlDelay -1
            if (p["classNN"]!="") {
                ControlClick p["classNN"], targetSpec, , "Left", 1, "NA"
                ShowClickFX_Mouse()
            } else {
                ControlClick "x" (p["x"]+dx) " y" (p["y"]+dy), targetSpec, , "Left", 1, "NA"
                ShowClickFX_ClientXY(targetHwnd, p["x"]+dx, p["y"]+dy)
            }
        } else {
            Click p["x"]+dx, p["y"]+dy
            ShowClickFX_ClientXY(targetHwnd, p["x"]+dx, p["y"]+dy)
            p["driftStep"] := Mod(p["driftStep"]+1, 4)
        }
        success := true
    } catch as e {
        success := false
    }

    if (IsObject(prevCtx))
        RestoreUserContext(prevCtx, targetHwnd)

    return success
}

; ----- Feedback visual ("●" por ~200ms)
ShowClickFX_Mouse() {
    CoordMode "Mouse", "Screen"
    MouseGetPos &sx, &sy
    ToolTip "●", sx+8, sy+8, 20
    SetTimer (() => ToolTip(, , , 20)), -200   ; run-once (período negativo). :contentReference[oaicite:3]{index=3}
    CoordMode "Mouse", "Client"
}

ShowClickFX_ClientXY(hwnd, cx, cy) {
    if (!hwnd) {
        ShowClickFX_Mouse()
        return
    }
    sx := cx, sy := cy
    if ClientToScreen(hwnd, &sx, &sy) {
        ToolTip "●", sx+6, sy+6, 20
        SetTimer (() => ToolTip(, , , 20)), -200
    }
}

ClientToScreen(hwnd, &x, &y) {
    try {
        pt := Buffer(8, 0)
        NumPut("Int", x, pt, 0), NumPut("Int", y, pt, 4)
        DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)  ; conversão cliente→ecrã via WinAPI. :contentReference[oaicite:4]{index=4}
        x := NumGet(pt, 0, "Int"), y := NumGet(pt, 4, "Int")
        return true
    } catch {
        return false
    }
}

; ----- Contexto do utilizador -----
SaveUserContext() {
    activeHwnd := WinExist("A")
    CoordMode "Mouse", "Screen"
    MouseGetPos &sx, &sy
    CoordMode "Mouse", "Client"
    return Map("hwnd", activeHwnd, "x", sx, "y", sy)
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
    if (p["exe"]!="")  s .= "ahk_exe " p["exe"] " "
    if (p["cls"]!="")  s .= "ahk_class " p["cls"] " "
    ; if (p["titlePart"]!="")  s .= p["titlePart"]
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
        lv.Add(, p["id"], p["title"], p["x"], p["y"], modo, p["classNN"], p["intervalSec"], "", p["activateBefore"]? "Sim":"Não", p["active"]? "Sim":"Não")
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
    countText := IniRead(iniFile, "Meta", "Count", "0"), count := IsNumber(countText) ? Integer(countText) : 0
    Loop count {
        sec := "Point." A_Index
        p := Map()
        p["id"] := SafeInt(IniRead(iniFile, sec, "id", A_Index))
        p["title"] := IniRead(iniFile, sec, "title", "")
        p["hwnd"] := SafeInt(IniRead(iniFile, sec, "hwnd", "0"))
        p["exe"] := IniRead(iniFile, sec, "exe", "")
        p["cls"] := IniRead(iniFile, sec, "cls", "")
        p["titlePart"] := IniRead(iniFile, sec, "titlePart", "")
        p["coordMode"] := IniRead(iniFile, sec, "coordMode", "Client")
        p["x"] := SafeInt(IniRead(iniFile, sec, "x", "0"))
        p["y"] := SafeInt(IniRead(iniFile, sec, "y", "0"))
        p["useCtrl"] := IniRead(iniFile, sec, "useCtrl", "0") = "1"
        p["classNN"] := IniRead(iniFile, sec, "classNN", "")

        secText := Trim(IniRead(iniFile, sec, "intervalSec", ""))
        if (IsNumber(secText))
            intervalSec := Max(1, Integer(secText))
        else {
            msText := Trim(IniRead(iniFile, sec, "intervalMs", ""))
            intervalSec := IsNumber(msText) ? Max(1, Ceil(Integer(msText)/1000.0)) : 1
        }
        p["intervalSec"] := intervalSec
        p["intervalMs"] := intervalSec * 1000

        p["activateBefore"] := IniRead(iniFile, sec, "activateBefore", "1") = "1"
        p["active"] := IniRead(iniFile, sec, "active", "1") = "1"
        p["nextDue"] := 0, p["lastClick"] := 0
        p["driftStep"] := 0
        p["jitterPhase"] := 0

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
        IniWrite p["intervalSec"], iniFile, sec, "intervalSec"
        IniWrite (p["activateBefore"]?1:0), iniFile, sec, "activateBefore"
        IniWrite (p["active"]?1:0), iniFile, sec, "active"
    }
}

SafeInt(val, def := 0) {
    val := Trim(val)
    return IsNumber(val) ? Integer(val) : def
}