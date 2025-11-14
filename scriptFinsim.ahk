; === SCRIPT DEFINITIVO Y CORREGIDO - AHK V2 ===
#Requires AutoHotkey v2.0
#SingleInstance
#Warn

; --- CONFIGURACIÓN ---
SetTitleMatchMode 1   ; 1 = el título EMPIEZA con el texto dado
Global TituloVentana := "ATM Simulator"
Global DesplegableBanco_ClassNN := "ComboBox1"
Global ListaTarjetas_ClassNN   := "ListBox1"
Global RutaIniSeleccionado := ""   ; ruta del .ini que vamos a usar

; --- Atajo de teclado para iniciar: Presiona Ctrl + Alt + T ---
^!t::{
    SeleccionarArchivoYFlujo()
}

; --- PASO 0: SELECCIONAR ARCHIVO, LEER GROUP/CARD Y PREGUNTAR QUÉ HACER ---
SeleccionarArchivoYFlujo() {
    global RutaIniSeleccionado

    rutaArchivo := FileSelect(, , "Selecciona el archivo .ini a usar", "Archivos INI (*.ini)")
    if (rutaArchivo = "")
        return  ; usuario canceló

    RutaIniSeleccionado := rutaArchivo

    ; Leer archivo para mostrar GROUP y CARD actuales
    try {
        contenido := FileRead(rutaArchivo)
    } catch as e {
        MsgBox("No se pudo leer el archivo seleccionado." . "`nDetalle: " e.Message, "Error")
        return
    }

    if (contenido = "") {
        MsgBox("El archivo seleccionado está vacío.", "Aviso")
        ; igual podemos seguir y cargarlo/ejecutarlo si querés
    }

    ; Buscar primeros valores de GROUP= y CARD=
    groupVal := ""
    cardVal  := ""

    delim := InStr(contenido, "`r`n") ? "`r`n" : "`n"
    lines := StrSplit(contenido, delim)

    for _, l in lines {
        if (groupVal = "") {
            if RegExMatch(l, "i)^\s*GROUP\s*=\s*(.*)$", &m) {
                groupVal := m[1]
            }
        }
        if (cardVal = "") {
            if RegExMatch(l, "i)^\s*CARD\s*=\s*(.*)$", &n) {
                cardVal := n[1]
            }
        }
        if (groupVal != "" && cardVal != "")
            break
    }

    texto := "Archivo seleccionado:" . "`n" rutaArchivo . "`n`n"
    texto .= "GROUP actual: " . (groupVal != "" ? groupVal : "(no encontrado)") . "`n"
    texto .= "CARD  actual: " . (cardVal  != "" ? cardVal  : "(no encontrado)") . "`n`n"
    texto .= "¿Deseás CAMBIAR estos valores según el banco y la tarjeta que selecciones?" . "`n"
    texto .= "(Si eliges 'No', se usará el archivo tal como está y se irá directo a reproducir el script.)"

    resp := MsgBox(texto, "Confirmar acción", "YesNo")

    if (resp = "No") {
        ; No modificamos el archivo, solo lo cargamos y damos Play
        CargarYReproducirScript()
        return
    }

    ; Si respondió Sí, seguimos con el flujo normal de cambio
    IniciarProceso()
}

; --- FUNCIÓN 1: INICIA EL PROCESO DE CAMBIO (BANCO/TARJETA) ---
IniciarProceso() {
    global TituloVentana, DesplegableBanco_ClassNN
    if !WinExist(TituloVentana) {
        MsgBox("La ventana '" TituloVentana "' no fue encontrada.", "Error")
        return
    }

    try {
        ArrayBancos := ControlGetItems(DesplegableBanco_ClassNN, TituloVentana)
    } catch as e {
        MsgBox("Error al leer la lista de bancos. Verifica el ClassNN." . "`nDetalle: " e.Message, "Error")
        return
    }

    if !ArrayBancos.Length {
        MsgBox("No se encontraron bancos en el desplegable.", "Aviso")
        return
    }

    ElegirBancoGUI(ArrayBancos)
}

; --- FUNCIÓN 2: GUI PARA ELEGIR UN BANCO ---
ElegirBancoGUI(ArrayBancos) {
    BancoGui := Gui()
    BancoGui.Title := "Paso 1: Selecciona un Banco"
    BancoGui.SetFont("s10", "Segoe UI")
    BancoGui.Add("Text",, "Selecciona el banco de la lista:")

    LV_Bancos := BancoGui.Add("ListView", "r15 w400 Sort", ["Bancos Disponibles"])
    LV_Bancos.OnEvent("DoubleClick", BancoSeleccionado)

    for _, banco in ArrayBancos {
        LV_Bancos.Add("", Trim(banco))
    }

    BancoGui.Add("Button", "Default w100", "OK").OnEvent("Click", BancoSeleccionado)
    BancoGui.Show()

    BancoSeleccionado(*) {
        fila := LV_Bancos.GetNext()
        if (fila = 0)
            return
        BancoElegido := LV_Bancos.GetText(fila)
        BancoGui.Destroy()
        ContinuarConBanco(BancoElegido)
    }
}

; --- FUNCIÓN 3: SELECCIONA EL BANCO EN LA APP Y LEE LAS TARJETAS ---
ContinuarConBanco(BancoElegido) {
    global TituloVentana, DesplegableBanco_ClassNN, ListaTarjetas_ClassNN

    WinActivate(TituloVentana)
    WinWaitActive(TituloVentana, , 2)

    try {
        ControlChooseString(BancoElegido, DesplegableBanco_ClassNN, TituloVentana)
    } catch as e {
        MsgBox("No se pudo seleccionar el banco '" BancoElegido "'." . "`nDetalle: " e.Message, "Error")
        return
    }

    Sleep(2000)

    try {
        ArrayTarjetas := ControlGetItems(ListaTarjetas_ClassNN, TituloVentana)
    } catch as e {
        MsgBox("Error al leer la lista de tarjetas. Verifica el ClassNN." . "`nDetalle: " e.Message, "Error")
        return
    }

    if !ArrayTarjetas.Length {
        MsgBox("No se encontraron tarjetas para este banco.", "Aviso")
        return
    }

    ElegirTarjetaGUI(BancoElegido, ArrayTarjetas)
}

; --- FUNCIÓN 4: GUI PARA ELEGIR TARJETA Y MODIFICAR EL INI ---
ElegirTarjetaGUI(BancoElegido, ArrayTarjetas) {
    TarjetaGui := Gui()
    TarjetaGui.Title := "Paso 2: Selecciona una Tarjeta"
    TarjetaGui.SetFont("s10", "Segoe UI")
    TarjetaGui.Add("Text",, "Selecciona la tarjeta para el banco: " BancoElegido)

    LV_Tarjetas := TarjetaGui.Add("ListView", "r15 w500", ["N° de Tarjeta"])
    LV_Tarjetas.OnEvent("DoubleClick", TarjetaSeleccionada)

    for _, tarjeta in ArrayTarjetas {
        LV_Tarjetas.Add("", Trim(tarjeta))
    }

    TarjetaGui.Add("Button", "Default w100", "OK").OnEvent("Click", TarjetaSeleccionada)
    TarjetaGui.Show()

    TarjetaSeleccionada(*) {
        fila := LV_Tarjetas.GetNext()
        if (fila = 0)
            return
        TarjetaElegida := LV_Tarjetas.GetText(fila)
        TarjetaGui.Destroy()
        ModificarArchivoIni(BancoElegido, TarjetaElegida)
    }
}

; --- FUNCIÓN 5: MODIFICA GROUP= Y CARD= EN EL ARCHIVO SELECCIONADO ---
ModificarArchivoIni(Grupo, Tarjeta) {
    global RutaIniSeleccionado

    rutaArchivo := RutaIniSeleccionado
    if (rutaArchivo = "") {
        MsgBox("No hay archivo .ini seleccionado en RutaIniSeleccionado.", "Error")
        return
    }

    try {
        original := FileRead(rutaArchivo)
        if (original = "")
            throw Error("El archivo está vacío o no se pudo leer.")

        ; Detectar separador original y si el archivo terminaba con salto
        delim    := InStr(original, "`r`n") ? "`r`n" : "`n"
        trailing := (StrLen(original) >= StrLen(delim)) && (SubStr(original, -StrLen(delim) + 1) = delim)

        lines := StrSplit(original, delim)

        totalGroup := 0, totalCard := 0

        for i, l in lines {
            ; GROUP=
            cG := 0
            l := RegExReplace(
                l
              , "i)^( *|\t*)(GROUP)(\s*)=.*$"
              , "$1$2$3=" Grupo
              , &cG
            )
            if (cG)
                totalGroup += cG

            ; CARD=
            cC := 0
            l := RegExReplace(
                l
              , "i)^( *|\t*)(CARD)(\s*)=.*$"
              , "$1$2$3=" Tarjeta
              , &cC
            )
            if (cC)
                totalCard += cC

            lines[i] := l
        }

        if (!totalGroup && !totalCard) {
            MsgBox("No se encontraron claves 'GROUP=' ni 'CARD=' en el archivo. No se hicieron cambios.", "Sin cambios")
            ; igual podemos seguir y reproducir el script
        } else {
            ; Reconstruir exactamente con el mismo separador y salto final si lo tenía
            out := ""
            for i, l in lines {
                if (i > 1)
                    out .= delim
                out .= l
            }
            if (trailing)
                out .= delim

            ; Backup y escritura sin BOM
            FileCopy(rutaArchivo, rutaArchivo ".bak", true)
            f := FileOpen(rutaArchivo, "w", "UTF-8-RAW")
            if (!f)
                throw Error("No se pudo abrir el archivo para escritura.")
            f.Write(out), f.Close()

            MsgBox(
                "Hecho: se actualizaron las apariciones." 
                . "`nGROUP = " Grupo "  (reemplazos: " totalGroup ")"
                . "`nCARD  = " Tarjeta "  (reemplazos: " totalCard ")"
            , "OK")
        }

        ; Después de modificar (o no), cargamos el script y damos Play
        CargarYReproducirScript()

    } catch as e {
        MsgBox("No se pudo modificar el archivo." . "`nDetalle: " e.Message, "Error")
    }
}

; --- FUNCIÓN 6: SCRIPTS -> PLAYBACK SCRIPT FROM -> THIS MACHINE... -> ABRIR EL .INI -> PLAY ---
CargarYReproducirScript() {
    global TituloVentana, RutaIniSeleccionado

    if (RutaIniSeleccionado = "") {
        MsgBox("No se dispone de la ruta del archivo .ini seleccionado.", "Error")
        return
    }

    ; Aseguramos foco en el ATM Simulator
    WinActivate(TituloVentana)
    if !WinWaitActive(TituloVentana, , 2) {
        MsgBox("No se pudo activar la ventana del ATM Simulator para cargar el script.", "Error")
        return
    }

    ; 1) Abrir menú Scripts (ALT + S)
    Send("!s")
    Sleep(250)

    ; 2) Elegir 'Playback Script From' (aceleradora: S)
    Send("s")
    Sleep(250)

    ; 3) En el submenú elegir 'This machine...' (aceleradora: T)
    Send("t")

    ; 4) Esperar el diálogo estándar de selección de archivo (Open/Abrir)
    if !WinWaitActive("ahk_class #32770", , 3) {
        MsgBox("No se detectó la ventana de selección de archivo (Open/Abrir).", "Error")
        return
    }

    ; 5) Pegar la RUTA COMPLETA en Edit1 y confirmar
    try {
        path := RutaIniSeleccionado
        path := Trim(path, '"')   ; por si viniera con comillas

        ControlFocus("Edit1", "ahk_class #32770")
        Sleep(150)
        Send("^a")
        Sleep(80)
        Send("{Del}")
        Sleep(80)
        SendText(path)
        Sleep(200)
        Send("{Enter}")           ; equivalente a hacer clic en 'Abrir'
    } catch as e {
        MsgBox("No se pudo completar la selección del archivo en el diálogo." . "`nDetalle: " e.Message, "Error")
        return
    }

    ; 6) Esperar la ventana "Play Back A Local Script File"
    playWin := "Play Back A Local Script File"
    if !WinWait(playWin, , 5) {
        MsgBox("No se encontró la ventana 'Play Back A Local Script File' para ejecutar el script.", "Error")
        return
    }
    WinActivate(playWin)
    WinWaitActive(playWin, , 2)

    ; 7) Intentar disparar el botón Play de varias formas
    try {
        ; a) Foco en Button12 y SPACE
        ControlFocus("Button12", playWin)
        Sleep(150)
        Send("{Space}")
        Sleep(300)

        ; b) Fallback: ALT+P (si el botón Play tiene P subrayada)
        ;   Esto no rompe nada si ya se ejecutó con SPACE.
        Send("!p")
    } catch as e {
        MsgBox("No se pudo activar el botón Play." . "`nDetalle: " e.Message, "Error")
    }
}
