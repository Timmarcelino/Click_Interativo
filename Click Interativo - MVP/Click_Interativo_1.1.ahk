#Requires AutoHotkey v2.0
#SingleInstance Force

; =========================
; Click Interativo — MVP v1.1
; =========================
; - "Configurar Ponto de Interação": escolher janela + capturar coordenadas + modo de clique + intervalo.
; - Vários Pontos de Interação (memória + INI).
; - "Ciclo de Interação": 1 SetTimer global (round-robin), cada ponto tem o seu intervalo.
; - Encerramento por data/hora (opcional) OU por ESC/fecho da janela alvo/erro.
; - Encerrar ciclo NÃO fecha o script.
; - NOVO: "Restaurar foco & cursor" por clique (apenas modo Coordenadas).

SendMode "Input"
CoordMode "Mouse", "Client"  ; Click/MouseGetPos ficam relativos à área cliente da janela ativa para o projeto. :contentReference[oaicite:1]{index=1}

; ---------- Persistência ----------
iniFile := A_ScriptDir "\Click_Interativo.ini"

; ---------- Estado ----------
global gPoints := []            ; {id,title,hwnd,exe,cls,titlePart,coordMode,x,y,useCtrl,classNN,intervalMs,activateBefore,active,nextDue,lastClick}
global gRunning := false, gPaused := false
global gTickMs := 100           ; período do scheduler (não é o intervalo de clique)
global gEndAt := ""             ; timestamp "YYYYMMDDHH24MISS" (vazio = sem fim programado)
global gCounters := {clicks: 0, startedAt: 0}

; ---------- GUI principal ----------
g := Gui("", "Click Interativo")     ; sem AlwaysOnTop
g.SetFont("s10")

; Barra de controlo
btnAdd     := g.Add("Button", "w190", "Configurar Ponto de Interação")
btnTest    := g.Add("Button", "x+6 w130", "Testar Ponto")
btnRemove  := g.Add("Button", "x+6 w130", "Remover Ponto")

; Lista de pontos
lv := g.Add("ListView", "xm w860 h260 Grid", ["ID","Janela/Aplicação","x","y","Modo","ClassNN","Intervalo(ms)","Ativar","Ativo"])
lv.ModifyCol(1, 40), lv.ModifyCol(2, 330), lv.ModifyCol(3, 60), lv.ModifyCol(4, 60)
lv.ModifyCol(5, 90), lv.ModifyCol(6, 150), lv.ModifyCol(7, 110), lv.ModifyCol(8, 70), lv.ModifyCol(9, 70)

; Linha de agendamento e flags
cbActivateDefault := g.Add("CheckBox", "xm y+6 Checked", "Ativar janela antes do clique (padrão ao criar Ponto)")
g.Add("Text", "x+12", "Encerrar em:")
dtEnd := g.Add("DateTime", "x+6 w200", "yyyy-MM-dd HH:mm")
cbUseEnd := g.Add("CheckBox", "x+6", "Usar data/hora")

; NOVO: restaurar contexto do utilizador (foco & cursor) após cada clique por coordenadas
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
    ; Handler robusto: função separada (OnEvent requer uma função/callback). :contentReference[oaicite:2]{index=2}
    btCap.OnEvent("Click", CapturePoint.Bind(winHwnds, ddl, edX, edY))

    w.Add("Text","xm y+10","Modo de clique:")
    rbCoords := w.Add("Radio", "Group Checked", "Coordenadas (Click)")
    rbCtrl   := w.Add("Radio", "x+12", "ControlClick")
    w.Add("Text","xm","ClassNN (opcional p/ ControlClick):")
    edClass  := w.Add("Edit", "w360", "")

    w.Add("Text","xm y+10","Intervalo por Ponto (ms):")
    edInterval := w.Add("Edit","w160 Number","1000")

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
        interval := Max(10, Integer(edInterval.Text))
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
        p["intervalMs"] := interval
        p["activateBefore"] := act
        p["active"] := true
        p["nextDue"] := 0, p["lastClick"] := 0

        gPoints.Push(p)
        SavePointsToIni()
        RefreshList()
    }

    w.Show()
}

; --- Handler separado para capturar coordenadas ---
CapturePoint(winHwnds, ddl, edX, edY, *) {
    idx := ddl.Value
    if (idx = 0) {
        MsgBox "Escolhe uma janela antes de capturar."
        return
    }
    ActivateWindow("ahk_id " winHwnds[idx])     ; ativa e aguarda ficar ativo (WinActivate/WinWaitActive). :contentReference[oaicite:3]{index=3}
    ToolTip "Clique e solte no ponto de interação..."
    KeyWait "LButton", "D"
    KeyWait "LButton", "U"
    ToolTip
    MouseGetPos &mx, &my                         ; lê posição atual (respeita CoordMode atual = Client). :contentReference[oaicite:4]{index=4}
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
    DoOneClick(p)  ; clique único segundo o modo do Ponto
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
    if (gRunning)
        return
    if (gPoints.Length=0) {
        MsgBox "Adicione pelo menos um Ponto."
        return
    }
    gRunning := true, gPaused := false
    gCounters := {clicks: 0, startedAt: A_TickCount}
    if (cbUseEnd.Value=1) {
        gEndAt := dtEnd.Value    ; 'YYYYMMDDHH24MISS' → comparar com A_Now
    } else {
        gEndAt := ""
    }
    now := A_TickCount
    for p in gPoints
        if (p["active"])
            p["nextDue"] := now

    btnStart.Enabled := false, btnStop.Enabled := true
    btnPause.Enabled := true, btnResume.Enabled := false
    lblStatus.Text := "Ciclo em execução…"
    SetTimer SchedulerTick, gTickMs             ; chama periodicamente mantendo GUI responsiva. :contentReference[oaicite:5]{index=5}
}

PauseCycle(*) {
    if (!gRunning || gPaused)
        return
    gPaused := true
    btnPause.Enabled := false, btnResume.Enabled := true
    lblStatus.Text := "Ciclo em pausa."
}

ResumeCycle(*) {
    if (!gRunning || !gPaused)
        return
    gPaused := false
    btnPause.Enabled := true, btnResume.Enabled := false
    lblStatus.Text := "Ciclo retomado."
}

StopCycle(*) {
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
        p["nextDue"] := now + p["intervalMs"]
        gCounters.clicks += 1
    }
}

; -------- Clique segundo o modo do ponto + restauração de contexto (quando aplicável)
DoOneClick(p) {
    ; Se for clique por coordenadas e o utilizador pediu restauração, salva contexto (janela ativa + cursor de ecrã)
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
        targetHwnd := WinExist(targetSpec)  ; pode ser 0 se janela não corresponder exatamente. :contentReference[oaicite:6]{index=6}

        if (p["useCtrl"]) {
            SetControlDelay -1                                  ; recomendação para ControlClick estável. :contentReference[oaicite:7]{index=7}
            if (p["classNN"]!="")
                ControlClick p["classNN"], targetSpec, , "Left", 1, "NA"
            else
                ControlClick "x" p["x"] " y" p["y"], targetSpec, , "Left", 1, "NA"
        } else {
            ; Clique por coordenadas: usa CoordMode "Client" (padrão do projeto) → Click respeita CoordMode. :contentReference[oaicite:8]{index=8}
            Click p["x"], p["y"]
        }
        success := true
    } catch as e {
        success := false
    }

    ; Restaura o contexto do utilizador (apenas se tínhamos capturado)
    if (IsObject(prevCtx))
        RestoreUserContext(prevCtx, targetHwnd)

    return success
}

; ----- Contexto do utilizador: salvar/recuperar janela ativa e cursor (em ecrã)
SaveUserContext() {
    ; Guarda HWND da janela ativa e cursor em SCREEN. WinExist("A") devolve HWND da ativa. :contentReference[oaicite:9]{index=9}
    activeHwnd := WinExist("A")
    CoordMode "Mouse", "Screen"
    MouseGetPos &sx, &sy                                   ; captura em coordenadas de ecrã. :contentReference[oaicite:10]{index=10}
    CoordMode "Mouse", "Client"                            ; volta ao modo do projeto
    ctx := Map()
    ctx["hwnd"] := activeHwnd
    ctx["x"] := sx, ctx["y"] := sy
    return ctx
}

RestoreUserContext(ctx, targetHwnd := 0) {
    ; Reativa a janela anterior (se existir e não for a mesma da interação) e repõe o cursor em SCREEN.
    try {
        if (ctx.Has("hwnd") && ctx["hwnd"] && ctx["hwnd"] != targetHwnd && WinExist("ahk_id " ctx["hwnd"])) {
            WinActivate "ahk_id " ctx["hwnd"]              ; ativa janela anterior. :contentReference[oaicite:11]{index=11}
            WinWaitActive "ahk_id " ctx["hwnd"], , 1       ; confirmação rápida. :contentReference[oaicite:12]{index=12}
        }
        CoordMode "Mouse", "Screen"
        MouseMove ctx["x"], ctx["y"], 0                    ; recoloca cursor exatamente onde estava.
        CoordMode "Mouse", "Client"
    } catch
        ; silêncio: melhor esforço
        return
}

; Assegura que a janela do ponto existe; se hwnd expirou, tenta por exe+class(+título parcial)
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

; Monta um WinTitle robusto: "ahk_exe X ahk_class Y" + (título parcial opcional)
BuildWinTitle(p) {
    s := ""
    if (p["exe"]!="")
        s .= "ahk_exe " p["exe"] " "
    if (p["cls"]!="")
        s .= "ahk_class " p["cls"] " "
    ; opcional: combinar com parte do título:
    ; if (p["titlePart"]!="")
    ;     s .= p["titlePart"]
    return RTrim(s)
}

; Ativa a janela e espera ficar ativa (com fallback leve)
ActivateWindow(winTitleOrHwnd, timeoutMs := 2000) {
    wt := IsInteger(winTitleOrHwnd) ? "ahk_id " winTitleOrHwnd : winTitleOrHwnd
    WinActivate wt                                     ; restaura se minimizada e ativa a janela. :contentReference[oaicite:13]{index=13}
    if WinWaitActive(wt, , timeoutMs/1000)             ; retorna HWND da ativa ou 0 no timeout. :contentReference[oaicite:14]{index=14}
        return true
    SendEvent "{Alt down}{Alt up}"                     ; “toque” para acordar z-order
    WinActivate wt
    return WinWaitActive(wt, , timeoutMs/1000) != 0
}

; ----- Lista / INI -----
RefreshList() {
    lv.Delete()
    for p in gPoints {
        modo := p["useCtrl"] ? "Control" : "Coords"
        lv.Add(, p["id"], p["title"], p["x"], p["y"], modo, p["classNN"], p["intervalMs"], p["activateBefore"]? "Sim":"Não", p["active"]? "Sim":"Não")
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
        p["intervalMs"] := Integer(IniRead(iniFile, sec, "intervalMs", "1000"))
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
        IniWrite p["intervalMs"], iniFile, sec, "intervalMs"
        IniWrite (p["activateBefore"]?1:0), iniFile, sec, "activateBefore"
        IniWrite (p["active"]?1:0), iniFile, sec, "active"
    }
}
