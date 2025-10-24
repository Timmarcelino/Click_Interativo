#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================
; Click Interativo — MVP v2.1.4 janela unívoca
; ============================================
; Novidade: identificação única da janela
; - Captura: salva hwnd, pid, título e tamanho do cliente (cw,ch)
; - Ativação: usa ahk_id quando possível; fallback com heurística
;   (PID+classe -> exe+classe) + (título/cliente mais próximos).
; Mantido: anti-concorrência, jitter ±s, drift ↑→↓←, feedback anel,
; redimensionável, INI, “Falta(s)”, ESC encerra.

SendMode "Input"
CoordMode "Mouse", "Client"

; ---------- Persistência ----------
iniFile := A_ScriptDir "\Click_Interativo.ini"

; ---------- Estado ----------
global gPoints := []   ; {id,title,hwnd,exe,cls,x,y,intervalSec,intervalMs,active,nextDue,lastClick,driftStep,jitterPhase,pid,cw,ch}
global gRunning := false, gPaused := false
global gTickMs := 100
global gEndAt := ""                    ; YYYYMMDDHH24MISS
global gCounters := {clicks: 0, startedAt: 0}
global gLastCountdownRefresh := 0
global gJitterSec := 1

; ----- Anti-concorrência -----
global gClickLock := false
global gLastGlobalClickMs := 0
global gMinInterClickMs := 60
global gStaggerMs := 50

; ----- Layout base (para redimensionamento) -----
global gLayout := Map()

; ---------- GUI principal ----------
g := Gui("+Resize", "Click Interativo")
g.SetFont("s10")

btnAdd     := g.Add("Button", "w200", "Configurar Ponto de Interação")
btnTest    := g.Add("Button", "x+6 w130", "Testar Ponto")
btnRemove  := g.Add("Button", "x+6 w130", "Remover Ponto")

; 1:ID 2:Janela 3:x 4:y 5:Intervalo(s) 6:Falta(s) 7:Ativo
lv := g.Add("ListView", "xm w820 h260 Grid", ["ID","Janela/Aplicação","x","y","Intervalo(s)","Falta(s)","Ativo"])
lv.ModifyCol(1, 40), lv.ModifyCol(2, 360), lv.ModifyCol(3, 60), lv.ModifyCol(4, 60)
lv.ModifyCol(5, 100), lv.ModifyCol(6, 90), lv.ModifyCol(7, 70)

; Agendamento
lblEnc := g.Add("Text", "xm y+6", "Encerrar em:")
dtDate := g.Add("DateTime", "x+6 w140", "yyyy-MM-dd")
dtTime := g.Add("DateTime", "x+6 w120", "HH:mm")
cbUseEnd := g.Add("CheckBox", "x+8", "Usar data/hora")

; Jitter
lblJit := g.Add("Text", "xm y+10", "Jitter (±s):")
edJitter := g.Add("Edit", "x+6 w60 Number", "1")

; Controlo do ciclo
btnStart  := g.Add("Button", "xm y+10 w150 Default", "Iniciar Ciclo")
btnPause  := g.Add("Button", "x+6 w120 Disabled", "Pausar")
btnResume := g.Add("Button", "x+6 w120 Disabled", "Retomar")
btnStop   := g.Add("Button", "x+6 w150 Disabled", "Encerrar Ciclo")
lblStatus := g.Add("Text", "xm y+8 w820", "Pronto.")
lblEsc    := g.Add("Text", "xm cGray", "Atalho: pressione ESC para encerrar o ciclo.")

; Eventos principais
g.OnEvent("Close", (*) => (gRunning ? StopCycle() : ExitApp()))
btnAdd.OnEvent("Click", AddPointWizard)
btnTest.OnEvent("Click", TestSelectedPoint)
btnRemove.OnEvent("Click", RemoveSelectedPoint)
btnStart.OnEvent("Click", StartCycle)
btnPause.OnEvent("Click", PauseCycle)
btnResume.OnEvent("Click", ResumeCycle)
btnStop.OnEvent("Click", StopCycle)
cbUseEnd.OnEvent("Click", (*) => ToggleEndInputs(cbUseEnd.Value=1))

; Hotkey: encerrar ciclo
Esc:: StopCycle()

; Inicialização
dtDate.Value := A_Now, dtTime.Value := A_Now
ToggleEndInputs(false)

; --- Mostrar, inicializar layout e registrar OnSize ---
g.Show()
InitLayout()
g.OnEvent("Size", Gui_OnSize)

LoadPointsFromIni()
RefreshList()

; ============================================
;              Implementação
; ============================================

InitLayout() {
    global gLayout, g
    g.GetPos(, , &w, &h)
    gLayout["baseW"] := w, gLayout["baseH"] := h

    lv.GetPos(&lvx,&lvy,&lvw,&lvh)
    gLayout["lvX"] := lvx, gLayout["lvY"] := lvy, gLayout["lvW"] := lvw, gLayout["lvH"] := lvh

    lblEnc.GetPos(, &encY), gLayout["encY"] := encY
    dtDate.GetPos(, &dtY),  gLayout["dtY"]  := dtY
    lblJit.GetPos(, &jitY), gLayout["jitY"] := jitY
    edJitter.GetPos(, &jitEdY), gLayout["jitEdY"] := jitEdY

    btnStart.GetPos(, &stY),   gLayout["stY"] := stY
    btnPause.GetPos(, &psY),   gLayout["psY"] := psY
    btnResume.GetPos(, &rsY),  gLayout["rsY"] := rsY
    btnStop.GetPos(, &spY),    gLayout["spY"] := spY

    lblStatus.GetPos(&lsX, &lsY, &lsW, &lsH)
    gLayout["lsX"] := lsX, gLayout["lsY"] := lsY, gLayout["lsW"] := lsW
    lblEsc.GetPos(, &escY), gLayout["escY"] := escY

    g.MinSize := w "x" h
}

Gui_OnSize(thisGui, MinMax, W, H) {
    if (MinMax = 1)
        return
    global gLayout
    if !(gLayout.Has("baseW"))
        return

    dw := W - gLayout["baseW"]
    dh := H - gLayout["baseH"]

    newLVW := Max(400, gLayout["lvW"] + dw)
    newLVH := Max(120, gLayout["lvH"] + dh)
    lv.Move(, , newLVW, newLVH)

    ny := gLayout["encY"] + dh, lblEnc.Move(, ny)
    ny := gLayout["dtY"]  + dh, dtDate.Move(, ny), dtTime.Move(, ny), cbUseEnd.Move(, ny)

    ny := gLayout["jitY"] + dh, lblJit.Move(, ny)
    ny := gLayout["jitEdY"] + dh, edJitter.Move(, ny)

    ny := gLayout["stY"] + dh, btnStart.Move(, ny)
    ny := gLayout["psY"] + dh, btnPause.Move(, ny)
    ny := gLayout["rsY"] + dh, btnResume.Move(, ny)
    ny := gLayout["spY"] + dh, btnStop.Move(, ny)

    ny := gLayout["lsY"] + dh
    newLSW := Max(300, gLayout["lsW"] + dw)
    lblStatus.Move(gLayout["lsX"], ny, newLSW)

    ny := gLayout["escY"] + dh, lblEsc.Move(, ny)
}

ToggleEndInputs(enable) {
    dtDate.Enabled := enable
    dtTime.Enabled := enable
}

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
        title := "", exe := "", cls := "", pid := 0
        try title := WinGetTitle("ahk_id " hwnd)
        try exe := WinGetProcessName("ahk_id " hwnd)
        try cls := WinGetClass("ahk_id " hwnd)
        try pid := WinGetPID("ahk_id " hwnd)

        x := Integer(edX.Text), y := Integer(edY.Text)
        if (x = "" || y = "") {
            MsgBox "Defina coordenadas válidas."
            return
        }
        secText := Trim(edInterval.Text)
        sec := (IsNumber(secText) ? Max(1, Integer(secText)) : 1)

        cw := 0, ch := 0
        GetClientSize(hwnd, &cw, &ch)

        p := Map()
        p["id"] := NewPointId()
        p["title"] := title
        p["hwnd"] := hwnd
        p["pid"] := pid
        p["exe"] := exe
        p["cls"] := cls
        p["cw"] := cw
        p["ch"] := ch
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

; --- Capturar coordenadas ---
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
    jt := Trim(edJitter.Text)
    gJitterSec := (IsNumber(jt) ? Abs(Integer(jt)) : 1)

    gRunning := true, gPaused := false
    gCounters := {clicks: 0, startedAt: A_TickCount}
    gClickLock := false, gLastGlobalClickMs := 0

    if (cbUseEnd.Value=1) {
        ds := dtDate.Value, ts := dtTime.Value
        gEndAt := SubStr(ds, 1, 8) . SubStr(ts, 9, 6)
    } else
        gEndAt := ""

    now := A_TickCount
    idx := 0
    for p in gPoints {
        if (p["active"]) {
            idx++
            p["nextDue"] := now + (idx-1)*gStaggerMs
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

    if (gClickLock || (now - gLastGlobalClickMs < gMinInterClickMs)) {
        UpdateCountdowns(now)
        return
    }

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
    prevCtx := SaveUserContext()
    success := false

    ; Ativa SEMPRE por ahk_id quando tivermos hwnd válido
    target := EnsureWindow(p) ? "ahk_id " p["hwnd"] : ""
    try {
        if (!target)
            throw Error("Sem janela alvo.")
        if !ActivateWindow(target)
            throw Error("Falha ao ativar janela alvo.")
        targetHwnd := WinExist(target)

        ; drift 1px no padrão ↑→↓←
        drift := [[0,-1],[1,0],[0,1],[-1,0]]
        step := Mod(p["driftStep"], 4)
        dx := drift[step+1][1], dy := drift[step+1][2]

        Click p["x"]+dx, p["y"]+dy
        ShowClickRipple_ClientXY(targetHwnd, p["x"]+dx, p["y"]+dy)

        p["driftStep"] := Mod(p["driftStep"]+1, 4)
        success := true
    } catch as e {
        success := false
    }
    RestoreUserContext(prevCtx)
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
RestoreUserContext(ctx) {
    try {
        if (ctx.Has("hwnd") && ctx["hwnd"] && WinExist("ahk_id " ctx["hwnd"])) {
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
    ; 1) Se hwnd válido: ok
    if (p.Has("hwnd") && p["hwnd"] && WinExist("ahk_id " p["hwnd"]))
        return true

    ; 2) Tenta encontrar por PID + classe
    cand := []
    if (p.Has("pid") && p["pid"]) {
        try {
            for hwnd in WinGetList("ahk_pid " p["pid"] " ahk_class " p["cls"])
                cand.Push(hwnd)
        }
    }

    ; 3) Se não achou, tenta por exe + classe
    if (cand.Length=0) {
        try {
            for hwnd in WinGetList("ahk_exe " p["exe"] " ahk_class " p["cls"])
                cand.Push(hwnd)
        }
    }

    if (cand.Length=0)
        return false

    ; 4) Escolhe a melhor candidata pelo título e tamanho do cliente
    best := PickBestWindow(cand, p)
    if (best) {
        p["hwnd"] := best
        return true
    }
    return false
}

PickBestWindow(candidates, p) {
    best := 0, bestScore := 10**9
    targetTitle := p.Has("title") ? p["title"] : ""
    tcw := p.Has("cw") ? p["cw"] : 0
    tch := p.Has("ch") ? p["ch"] : 0

    for hwnd in candidates {
        t := ""
        try t := WinGetTitle("ahk_id " hwnd)

        cw := 0, ch := 0
        GetClientSize(hwnd, &cw, &ch)

        ; score: menor é melhor
        ; título: igual (0), contém (10), diferente (30)
        titleScore := 30
        if (targetTitle != "") {
            if (t = targetTitle)
                titleScore := 0
            else if InStr(t, targetTitle)
                titleScore := 10
        }

        ; tamanho do cliente: distância Manhattan
        sizeScore := (tcw>0 && tch>0) ? Abs(cw - tcw) + Abs(ch - tch) : 50

        s := titleScore*1000 + sizeScore  ; título pesa mais
        if (s < bestScore) {
            bestScore := s
            best := hwnd
        }
    }
    return best
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
        p["pid"] := SafeInt(IniRead(iniFile, sec, "pid", "0"))
        p["exe"] := IniRead(iniFile, sec, "exe", "")
        p["cls"] := IniRead(iniFile, sec, "cls", "")
        p["cw"] := SafeInt(IniRead(iniFile, sec, "cw", "0"))
        p["ch"] := SafeInt(IniRead(iniFile, sec, "ch", "0"))
        p["x"] := SafeInt(IniRead(iniFile, sec, "x", "0"))
        p["y"] := SafeInt(IniRead(iniFile, sec, "y", "0"))

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
        IniWrite p["pid"], iniFile, sec, "pid"
        IniWrite p["exe"], iniFile, sec, "exe"
        IniWrite p["cls"], iniFile, sec, "cls"
        IniWrite p["cw"], iniFile, sec, "cw"
        IniWrite p["ch"], iniFile, sec, "ch"
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

; ----- WinAPI helpers -----
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

GetClientSize(hwnd, &w, &h) {
    w := 0, h := 0
    try {
        rect := Buffer(16, 0) ; RECT {left,top,right,bottom}
        if DllCall("GetClientRect", "ptr", hwnd, "ptr", rect) {
            w := NumGet(rect, 8, "Int")
            h := NumGet(rect, 12, "Int")
            return true
        }
    }
    return false
}