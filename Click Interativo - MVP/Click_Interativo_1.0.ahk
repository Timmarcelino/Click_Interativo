#Requires AutoHotkey v2.0
#SingleInstance Force

; =========================
; Clicker Interactivo — v1 (MVP)
; =========================
; - "Configurar Ponto de Interação": escolher janela + capturar coordenadas + modo de clique + intervalo.
; - Vários Pontos de Interação (armazenados em memória e INI).
; - "Ciclo de Interação": 1 SetTimer global (round-robin), cada ponto tem o seu intervalo.
; - Encerramento por data/hora (opcional) OU por ESC/fecho da janela alvo/erro.
; - Encerrar ciclo NÃO fecha o script.

SendMode "Input"
CoordMode "Mouse", "Client"  ; Click/MouseGetPos passam a ser relativos à área cliente da janela ativa. (Doc: CoordMode) 
