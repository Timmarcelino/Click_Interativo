# Click Interactivo (AutoHotkey v2)

Automatiza cliques em aplicações Windows com controlo fino. O núcleo do projeto são **Pontos de Interação** (janela + coordenadas, opcionalmente `ClassNN`) e um **Ciclo de Interação** que dispara cliques nesses pontos por um **agendador único** (um `SetTimer` global). O ciclo pode ser pausado/retomado/encerrado e, opcionalmente, termina numa data/hora definida. O encerramento do ciclo **não** fecha o script.

---

## Visão geral do projeto

- **Ponto de Interação** = *Janela/Aplicação* + *Coordenadas do clique* (ou *ControlClick* com `ClassNN`).  
  Coordenadas obedecem ao `CoordMode` (por padrão, **Client**: relativas à área cliente da janela ativa).
- **Configurar Ponto de Interação**: assistente para selecionar a janela, ativá‑la, capturar `X,Y` (no *mouse‑up*) e definir modo/intervalo por ponto.
- **Ciclo de Interação**: um único `SetTimer` percorre os pontos e dispara cada um **quando vence o seu próprio intervalo** (`now >= nextDue`), mantendo a GUI responsiva.
- **Clique**:
  - `Click x,y` (respeita o `CoordMode`).
  - `ControlClick` (por `ClassNN` ou `"xN yM"`), com `SetControlDelay -1` e opção `NA` para máxima fiabilidade sem “roubar” foco.

---

## Arquitetura e componentes principais

- **GUI principal** (Painel de controlo)  
  Lista de pontos (criar/testar/remover), *Ativar antes de clicar* (flag), planeamento de **Data/Hora de fim** e comandos *Iniciar, Pausar, Retomar, Encerrar*.
- **Modelo de dados**
  ```txt
  InteractionPoint {
    id, title, x, y, coordMode,
    useCtrl, classNN, intervalMs,
    activateBefore, exe, cls, titlePart,
    hwnd (volátil), nextDue, lastClick, active
  }
  ```
  Identificação robusta da janela por `ahk_exe`/`ahk_class` (e opcionalmente parte do título), além do `HWND` corrente.
- **Motor de clique**  
  Ativa janela (quando configurado) com `WinActivate` e confirma com `WinWaitActive` (fallback leve com *toque de Alt*); depois dispara `Click` ou `ControlClick`.

---

## Sistema de coordenadas e configurações

- **Padrão**: `CoordMode "Mouse","Client"` — coordenadas relativas à **área cliente da janela ativa** (mais estável em DPI/ecrãs múltiplos). `Click`, `MouseMove` e `MouseGetPos` respeitam este modo.
- **Captura**: ativar janela → aguardar *mouse‑up* → `MouseGetPos(&x,&y)` → guardar `x,y`.
- **Alternativa estável**: `ControlClick` com `ClassNN` quando possível; usar `SetControlDelay -1` e `"NA"` para não interferir com o ponteiro real.

---

## Fluxo de dados e armazenamento

- **Carregamento**: ficheiro `Click_Interactivo.ini` lido via `IniRead`, secções `[Point.N]`.
- **Persistência**: alterações escritas com `IniWrite`. INI é simples, portátil e auditável.

Exemplo de estrutura INI:
```ini
[Meta]
Count=2

[Point.1]
id=1
title=Bloco de Notas
exe=notepad.exe
cls=Notepad
coordMode=Client
x=120
y=80
useCtrl=0
classNN=
intervalMs=700
activateBefore=1
active=1
```

---

## Ciclo de execução e condições de encerramento

- **Agendamento**: `SetTimer SchedulerTick, 100` (tick global). Cada ponto dispara quando `A_TickCount >= nextDue`.
- **Ativação condicional**:  
  - `Click` por coordenadas → recomenda‑se ativar (`WinActivate` + `WinWaitActive`).  
  - `ControlClick` com `ClassNN` + `"NA"` → ativação tende a ser desnecessária.
- **Encerramento do ciclo** (o script **permanece aberto**):
  - Manual: botão *Encerrar* ou tecla **Esc**.
  - **Janela ausente**: `WinExist` falha → parar ciclo com estado.
  - **Data/hora final**: `DateTime` GUI → timestamp `YYYYMMDDHH24MISS`; comparar com `A_Now`.

---

## Convenções de código

- **AutoHotkey v2**: usar funções/classes pequenas; `try/catch` em operações de janela/controlo.
- **Nomenclatura**: evitar colisão entre nomes de classe e variáveis (identificadores não diferenciam maiúsculas/minúsculas).
- **GUI**: `OnEvent` para *handlers*; evitar modais dentro de `SchedulerTick`.
- **WinTitle**: combinar `ahk_exe` e `ahk_class` (e opcional `titlePart`) para reanexar janelas em execuções futuras.

---

## Dependências

- **AutoHotkey v2** instalado.
- **Permissões Windows**: interagir com apps elevadas pode exigir executar o script como administrador.
- **(Opcional)**: bibliotecas utilitárias de data/hora se precisares de parsing extra (o projeto usa `DateTime` nativo da GUI).

---

## Ficheiros chave

- `Click_Interactivo.ahk` — fonte principal (GUI, modelos, agendador, motor de clique).
- `Click_Interactivo.ini` — persistência dos Pontos de Interação (um por secção).

---

## Observações importantes para desenvolvimento

- **Preferir `ControlClick`** com `ClassNN` quando o alvo for um controlo estável; caso contrário, `Click` + `CoordMode "Client"`. Use `SetControlDelay -1` e `"NA"` para máxima fiabilidade.
- **Ativação consciente**: confirmar foco antes de `Click` por coordenadas com `WinWaitActive`.
- **Identificação de janela**: além do `HWND` (volátil), persistir `exe`/`cls` e (se fizer sentido) parte do título.

---

## Roadmap

1. **MVP**: criar/testar/remover Pontos; ciclo com 1 `SetTimer`; pausa/retoma/encerramento; data/hora final; hotkey **Esc**.
2. **Edição de Pontos**: alterar intervalo/modo/`ClassNN`/ativação; “Ir para janela alvo”.
3. **Matching de janela**: usar `titlePart` opcional e ordem de *matching* `HWND → exe+class → título`.
4. **Paragem por evento (opcional)**: `PixelSearch`/`ImageSearch` para encerrar/pular clique mediante mudança visual.
5. **Qualidade de vida**: duplicar/reordenar pontos; exportar/importar INI; logs.
6. **Modo avançado**: jitter de intervalos; múltiplos timers (um por ponto) com trechos críticos mínimos, se a precisão individual por ponto se tornar requisito.
