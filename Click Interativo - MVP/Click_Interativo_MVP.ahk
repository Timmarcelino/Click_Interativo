#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================
; Click Interativo — MVP v2.1.1 
; ============================================
; Mantém: Data/Hora de término (opcional), contador “Falta(s)”,
; drift (↑→↓←), jitter ±s alternado, feedback anel pulsante,
; SetTimer único, persistência INI, intervalos por ponto (segundos),
; ESC encerra ciclo.
; Novo: garantia de 1 clique por vez (lock + lacuna global + stagger).

SendMode "Input"
CoordMode "Mouse", "Client"      ; Click/MouseGetPos relativos à área cliente

; ---------- Persistência ----------
iniFile := A_ScriptDir "\Click_Interativo.ini"

; ---------- Estado ----------
global gPoints := []   ; {id,title,hwnd,exe,cls,x,y,intervalSec,intervalMs,active,nextDue,lastClick,driftStep,jitterPhase}
global gRunning := false, gPaused := false
global gTickMs := 100
global gEndAt := ""                    ; YYYYMMDDHH24MISS
global gCounters := {clicks: 0, startedAt: 0}
global gLastCountdownRefresh := 0
global gJitterSec := 1

; ----- Anti-concorrência -----
global gClickLock := false             ; true enquanto um clique está em execução
global gLastGlobalClickMs := 0
global gMinInterClickMs := 60          ; lacuna mínima entre cliques (ms)
global gStaggerMs := 50                ; offset inicial por ponto ao iniciar o ciclo (ms)

; ---------- GUI principal ----------
g := Gui("", "Click Interativo")
g.SetFont("s10")

btnAdd     := g.Add("Button", "w200", "Configurar Ponto de Interação")
btnTest    := g.Add("Button", "x+6 w130", "Testar Ponto")
btnRemove  := g.Add("Button", "x+6 w130", "Remover Ponto")

; 1:ID 2:Janela 3:x 4:y 5:Intervalo(s) 6:Falta(s) 7:Ativo
lv := g.Add("ListView", "xm w820 h260 Grid", ["ID","Janela/Aplicação","x","y","Intervalo(s)","Falta(s)","Ativo"])
lv.ModifyCol(1, 40), lv.ModifyCol(2, 360), lv.ModifyCol(3, 60), lv.ModifyCol(4, 60)
lv.ModifyCol(5, 100), lv.ModifyCol(6, 90), lv.ModifyCol(7, 70)

; Agendamento
g.Add("Text", "xm y+6", "Encerrar em:")
dtDate := g.Add("DateTime", "x+6 w140", "yyyy-MM-dd")
dtTime := g.Add("DateTime", "x+6 w120", "HH:mm")
cbUseEnd := g.Add("CheckBox", "x+8", "Usar data/hora")

; Jitter
g.Add("Text", "xm y+10", "Jitter (±s):")
edJitter := g.Add("Edit", "x+6 w60 Number", "1")

; Controlo do ciclo
btnStart  := g.Add("Button", "xm y+10 w150 Default", "Iniciar Ciclo")
btnPause  := g.Add("Button", "x+6 w120 Disabled", "Pausar")
btnResume := g.Add("Button", "x+6 w120 Disabled", "Retomar")
btnStop   := g.Add("Button", "x+6 w150 Disabled", "Encerrar Ciclo")
lblStatus := g.Add("Text", "xm y+8 w820", "Pronto.")
g.Add("Text", "xm cGray", "Atalho: pressione ESC para encerrar o ciclo.")

; Ligações de eventos
g.OnEvent("Close", (*) => (gRunning ? StopCycle() : ExitApp()))
btnAdd.OnEvent("Click", AddPointWizard)
btnTest.OnEvent("Click", TestSelectedPoint)
btnRemove.OnEvent("Click", RemoveSelectedPoint)
btnStart.OnEvent("Click", StartCycle)
btnPause.OnEvent("Click", PauseCycle)
btnResume.OnEvent("Click", ResumeCycle)
btnStop.OnEvent("Click", StopCycle)
cbUseEnd.OnEvent("Click", (*) => ToggleEndInputs(cbUseEnd.Value=1))

; Hotkey: encerrar ciclo (não fecha script)
Esc:: StopCycle()

; Inicialização de valores nos DateTime e estado desativado
dtDate.Value := A_Now, dtTime.Value := A_Now
ToggleEndInputs(false)

g.Show()
LoadPointsFromIni()
RefreshList()

; ============================================
;              Implementação
; ============================================

ToggleEndInputs(enable) {
    dtDate.Enabled := enable
    dtTime.Enabled := enable
}

; ----- Wizard: Configurar Ponto de Interação (apenas coordenadas) -----
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

    w.Add("Text","xm y+10","Intervalo por Ponto (segundos):")
    edInterval := w.Add("Edit","w160 Number","1")

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
        p["x"] := x, p["y"] := y
        p["intervalSec"] := sec
        p["intervalMs"] := sec * 1000
        p["active"] := true
        p["nextDue"] := 0, p["lastClick"] := 0
        p["driftStep"] := 0
        p["jitterPhase"] := 0

        gPoints.Push(p), SavePointsToIni(), RefreshList()
    }

    w.Show()
}

; --- Capturar coordenadas (pressione e solte o botão esquerdo) ---
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
    global gRunning, gPaused, gCounters, gEndAt, gPoints, btnStart, btnStop, btnPause, btnResume, lblStatus, dtDate, dtTime, cbUseEnd, edJitter, gJitterSec
    global gClickLock, gLastGlobalClickMs

    if (gRunning)
        return
    if (gPoints.Length=0) {
        MsgBox "Adicione pelo menos um Ponto."
        return
    }
    ; jitter global
    jt := Trim(edJitter.Text)
    gJitterSec := (IsNumber(jt) ? Abs(Integer(jt)) : 1)

    gRunning := true, gPaused := false
    gCounters := {clicks: 0, startedAt: A_TickCount}
    gClickLock := false, gLastGlobalClickMs := 0

    if (cbUseEnd.Value=1) {
        ; DateTime.Value retorna YYYYMMDDHH24MISS (data e hora); juntamos data do 1º e hora do 2º
        ds := dtDate.Value, ts := dtTime.Value
        gEndAt := SubStr(ds, 1, 8) . SubStr(ts, 9, 6)    ; YYYYMMDD + HHMISS
    } else
        gEndAt := ""

    now := A_TickCount
    idx := 0
    for p in gPoints {
        if (p["active"]) {
            idx++
            p["nextDue"] := now + (idx-1)*gStaggerMs   ; stagger inicial p/ evitar simultaneidade
            p["driftStep"] := 0
            p["jitterPhase"] := 0
        }
    }

    btnStart.Enabled := false, btnStop.Enabled := true
    btnPause.Enabled := true, btnResume.Enabled := false
    lblStatus.Text := "Ciclo em execução…  (ESC encerra)"
    SetTimer SchedulerTick, gTickMs
}

PauseCycle(*) {
    global gRunning, gPaused, btnPause, btnResume, lblStatus
    if (!gRunning || gPaused)
        return
    gPaused := true
    btnPause.Enabled := false, btnResume.Enabled := true
    lblStatus.Text := "Ciclo em pausa.  (ESC encerra)"
}

ResumeCycle(*) {
    global gRunning, gPaused, btnPause, btnResume, lblStatus
    if (!gRunning || !gPaused)
        return
    gPaused := false
    btnPause.Enabled := true, btnResume.Enabled := false
    lblStatus.Text := "Ciclo retomado.  (ESC encerra)"
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
    global gClickLock, gLastGlobalClickMs, gMinInterClickMs, gJitterSec

    if (!gRunning || gPaused)
        return

    if (gEndAt != "" && A_Now >= gEndAt) {
        StopCycle()
        return
    }

    now := A_TickCount

    ; respeita lacuna mínima e não reentra
    if (gClickLock || (now - gLastGlobalClickMs < gMinInterClickMs)) {
        UpdateCountdowns(now)
        return
    }

    ; escolher UM ponto devido: o de menor nextDue
    bestIdx := 0, bestDue := 0x7FFFFFFF
    idx := 0
    for p in gPoints {
        idx++
        if (!p["active"])
            continue
        if (now >= p["nextDue"] && p["nextDue"] < bestDue) {
            bestDue := p["nextDue"]
            bestIdx := idx
        }
    }

    if (bestIdx = 0) {
        UpdateCountdowns(now)
        return
    }

    ; bloquear e executar um único clique
    gClickLock := true
    p := gPoints[bestIdx]

    if !EnsureWindow(p) {
        lblStatus.Text := "Janela alvo ausente. Encerrando ciclo."
        gClickLock := false
        StopCycle()
        return
    }

    ok := DoOneClick(p)
    if (!ok) {
        lblStatus.Text := "Erro no clique. Encerrando ciclo."
        gClickLock := false
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
    lblStatus.Text := "Clique em: " p["title"] "  (" gCounters.clicks " total)  —  ESC encerra"

    gLastGlobalClickMs := A_TickCount
    gClickLock := false

    ; atualizar contadores no fim do tick
    UpdateCountdowns(A_TickCount)
}

UpdateCountdowns(now) {
    global gLastCountdownRefresh, gPoints, lv
    if (now - gLastCountdownRefresh >= 250) {
        gLastCountdownRefresh := now
        idx := 0
        for p in gPoints {
            idx++
            if (!p["active"]) {
                lv.Modify(idx, "Col6", "")
                continue
            }
            rem := Max(0, p["nextDue"] - now) / 1000.0
            lv.Modify(idx, "Col6", Format("{:.1f}", rem))
        }
    }
}

; -------- Clique (coords) + drift + feedback + restauração de contexto
DoOneClick(p) {
    ; salvar contexto do utilizador SEMPRE (janela ativa + mouse em ecrã)
    prevCtx := SaveUserContext()

    success := false
    targetSpec := BuildWinTitle(p)
    targetHwnd := 0

    ; micro-deslocamento (↑,→,↓,←)
    drift := [[0,-1],[1,0],[0,1],[-1,0]]
    step := Mod(p["driftStep"], 4)
    dx := drift[step+1][1], dy := drift[step+1][2]

    try {
        ; ativar sempre a janela alvo
        if !ActivateWindow(targetSpec)
            throw Error("Falha ao ativar janela alvo.")
        targetHwnd := WinExist(targetSpec)

        ; clique por coordenadas (respeita CoordMode "Client")
        Click p["x"]+dx, p["y"]+dy

        ; feedback visual: anel pulsante em COORDENADAS DE ECRÃ
        ShowClickRipple_ClientXY(targetHwnd, p["x"]+dx, p["y"]+dy)

        ; próximo passo do drift
        p["driftStep"] := Mod(p["driftStep"]+1, 4)
        success := true
    } catch as e {
        success := false
    }

    ; restaurar janela/cursor do utilizador
    RestoreUserContext(prevCtx, targetHwnd)
    return success
}

; =========================
;  FEEDBACK VISUAL (anel)
; =========================
ShowClickRipple_ClientXY(hwnd, cx, cy) {
    if (!hwnd) {
        CoordMode "Mouse", "Screen"
        MouseGetPos &sx, &sy
        CoordMode "Mouse", "Client"
        ShowClickRipple_Screen(sx, sy)
        return
    }
    sx := cx, sy := cy
    if ClientToScreen(hwnd, &sx, &sy) {
        ShowClickRipple_Screen(sx, sy)
    }
}

ShowClickRipple_Screen(sx, sy) {
    try {
        fx := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 -DPIScale")
        fx.BackColor := "Red"
        frames := [ Map("r",12,"a",230)
                  , Map("r",18,"a",170)
                  , Map("r",24,"a",110)
                  , Map("r",30,"a",70) ]
        idx := 1
        showFrame := (*) => (
            idx > frames.Length
                ? ( SetTimer(showFrame, 0), TryDestroyGui(fx) )
                : ( DrawRippleFrame(fx, sx, sy, frames[idx]["r"], frames[idx]["a"]), idx++ )
        )
        showFrame()
        SetTimer showFrame, 50
    } catch {
        ToolTip "●", sx+6, sy+6, 20
        SetTimer (() => ToolTip(, , , 20)), -250
    }
}

DrawRippleFrame(fx, sx, sy, r, alpha) {
    w := r*2, h := w
    fx.Show("NA x" (sx - r) " y" (sy - r) " w" w " h" h)
    WinSetTransparent alpha, fx.Hwnd
    WinSetRegion "0-0 w" w " h" h " E", fx.Hwnd
}

TryDestroyGui(fx) {
    try fx.Destroy()
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
        lv.Add(, p["id"], p["title"], p["x"], p["y"], p["intervalSec"], "", p["active"]? "Sim":"Não")
    }
}

NewPointId() {
    max := 0
    for p in gPoints
        if (p["id"]>max)
            max := p["id"]
    return max+1
}

; Parser robusto do INI (compatível com versões anteriores)
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
        p["x"] := SafeInt(IniRead(iniFile, sec, "x", "0"))
        p["y"] := SafeInt(IniRead(iniFile, sec, "y", "0"))

        ; intervalo em segundos com fallback legado (intervalMs)
        secText := Trim(IniRead(iniFile, sec, "intervalSec", ""))
        if (IsNumber(secText))
            intervalSec := Max(1, Integer(secText))
        else {
            msText := Trim(IniRead(iniFile, sec, "intervalMs", ""))
            intervalSec := IsNumber(msText) ? Max(1, Ceil(Integer(msText)/1000.0)) : 1
        }
        p["intervalSec"] := intervalSec
        p["intervalMs"] := intervalSec * 1000

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
        IniWrite p["x"], iniFile, sec, "x"
        IniWrite p["y"], iniFile, sec, "y"
        IniWrite p["intervalSec"], iniFile, sec, "intervalSec"
        IniWrite (p["active"]?1:0), iniFile, sec, "active"
    }
}

SafeInt(val, def := 0) {
    val := Trim(val)
    return IsNumber(val) ? Integer(val) : def
}

; ----- Conversão cliente→ecrã (WinAPI) -----
ClientToScreen(hwnd, &x, &y) {
    try {
        pt := Buffer(8, 0)
        NumPut("Int", x, pt, 0), NumPut("Int", y, pt, 4)
        DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
        x := NumGet(pt, 0, "Int"), y := NumGet(pt, 4, "Int")
        return true
    } catch {
        return false
    }
}
