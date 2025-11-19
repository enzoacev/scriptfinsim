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

; --------------------------------------------------------------
; FUNCIÓN AUXILIAR: NORMALIZA SALTOS DE LÍNEA A `r`n
; --------------------------------------------------------------
NormalizarSaltos(texto) {
    if (texto = "")
        return ""
    tmp := StrReplace(texto, "`r`n", "`n")
    tmp := StrReplace(tmp, "`r", "`n")
    return StrReplace(tmp, "`n", "`r`n")
}

; --------------------------------------------------------------
; PASO 0:
;   - LISTAR TRANSACCIONES BASE: archivos 1- NombreDeTransaccion*.ini en \transacciones
;   - SELECCIÓN CON CHECKBOXES
;   - SIEMPRE CREAR NUEVO SCRIPT COMBINADO
;   - NOMBRE POR DEFECTO: DDMMAAAA-HHMM-script.ini
;   - GUARDAR EN: \transacciones\scriptsCreados
;   - BOTÓN PARA USAR UN SCRIPT YA CREADO (de \scriptsCreados)
; --------------------------------------------------------------
SeleccionarArchivoYFlujo() {
    global RutaIniSeleccionado

    baseDir    := A_ScriptDir                ; carpeta donde está el script / exe
    transDir   := baseDir "\transacciones"   ; subcarpeta con transacciones base
    scriptsDir := transDir "\scriptsCreados" ; subcarpeta para scripts generados

    if !DirExist(transDir) {
        MsgBox("No se encontró la carpeta 'transacciones' en:" . "`n" transDir, "Error")
        return
    }

    if !DirExist(scriptsDir) {
        DirCreate(scriptsDir)
    }

    ; --- MOVER CUALQUIER SCRIPT GENERADO SUELTO AL SUBDIRECTORIO scriptsCreados ---
    Loop Files, transDir "\*.ini" {
        nombre := A_LoopFileName
        if RegExMatch(nombre, "^\d{8}-\d{4}-script\.ini$") {
            ; Lo mandamos a scriptsCreados
            FileMove(A_LoopFileFullPath, scriptsDir "\" nombre, 1)
        }
    }

    ; --- CARGAR SOLO TRANSACCIONES BASE (N- NombreDeTransaccion...) ---
    archivos := []  ; {nombre, ruta}
    Loop Files, transDir "\*.ini" {
        nombre := A_LoopFileName
        ruta   := A_LoopFileFullPath

        ; Solo los que empiezan con "N- " (1-, 2-, 3-, etc.)
        ; Ej: "1- ConsultaDeSaldo.ini", "2- ExtraccionPesosCA.ini"
        ; NO matchea "19112025-1558-script.ini" porque ahí no hay espacio después del guion
        if RegExMatch(nombre, "i)^\s*\d+\s*-\s+(.+)$") {
            archivos.Push({nombre: nombre, ruta: ruta})
        }
    }


; Ordenar alfabéticamente por nombre (método manual, sin .Sort)
for i, _ in archivos {
    if (i >= archivos.Length)
        break
    j := i + 1
    while (j <= archivos.Length) {
        a := archivos[i]
        b := archivos[j]
        ; StrCompare > 0 => a.nombre va "después" de b.nombre
        if (StrCompare(a.nombre, b.nombre) > 0) {
            temp := archivos[i]
            archivos[i] := archivos[j]
            archivos[j] := temp
        }
        j++
    }
}


    TransGui := Gui()
    TransGui.Title := "Transacciones base y scripts creados"
    TransGui.SetFont("s10", "Segoe UI")

    TransGui.Add("Text",, "Base de transacciones. Marcá una o varias:")

    LV_Arch := TransGui.Add("ListView", "r15 w520 Checked Grid", ["Archivo"])
    for idx, info in archivos {
        LV_Arch.Add("", info.nombre)
    }

    btnCombinar     := TransGui.Add("Button", "y+10 w150 Default", "Crear script nuevo")
    btnSelectAll    := TransGui.Add("Button", "x+10 w130", "Seleccionar todo")
    btnScriptPrevio := TransGui.Add("Button", "x+10 w140", "Usar script creado")
    btnCancel       := TransGui.Add("Button", "x+10 w90", "Cancelar")

    btnCombinar.OnEvent("Click", ConfirmarSeleccionBase)
    btnSelectAll.OnEvent("Click", SeleccionarTodo)
    btnScriptPrevio.OnEvent("Click", AbrirScriptsCreados)
    btnCancel.OnEvent("Click", (*) => TransGui.Destroy())
    LV_Arch.OnEvent("DoubleClick", ToggleCheckboxODisparar)

    TransGui.Show()

    ; --- Función local: seleccionar todo ---
    SeleccionarTodo(*) {
        total := LV_Arch.GetCount()
        if (total = 0)
            return

        ; Contar cuántas están checkeadas
        checked := 0
        row := 0
        while (row := LV_Arch.GetNext(row, "C")) {  ; "C" = Checked
            checked++
        }

        if (checked = total) {
            ; Todas estaban seleccionadas -> deseleccionar todas
            Loop total {
                LV_Arch.Modify(A_Index, "-Check")
            }
        } else {
            ; No todas estaban seleccionadas -> seleccionar todas
            Loop total {
                LV_Arch.Modify(A_Index, "Check")
            }
        }
    }



    ; --- Función local: cuando hacés doble click en una fila ---
    ToggleCheckboxODisparar(*) {
        fila := LV_Arch.GetNext()
        if (fila = 0)
            return
        ; Toggle de check
        estado := LV_Arch.GetChecked(fila)
        LV_Arch.Modify(fila, estado ? "-Check" : "Check")
    }

    ; --- Función local: combinar transacciones base seleccionadas ---
    ConfirmarSeleccionBase(*) {
        if (archivos.Length = 0) {
            MsgBox("No hay transacciones base (1- NombreDeTransaccion) en la carpeta.", "Aviso")
            return
        }

        seleccionados := []
        row := 0
        ; "C" -> filas CHECKED
        while (row := LV_Arch.GetNext(row, "C")) {
            seleccionados.Push(row)
        }

        if (seleccionados.Length = 0) {
            MsgBox("No seleccionaste ninguna transacción base.", "Aviso")
            return
        }

        ; Construir contenido nuevo combinando las seleccionadas
        delim           := "`r`n"
        contenidoNuevo  := ""
        primeraVez      := true

        for _, rowIdx in seleccionados {
            info := archivos[rowIdx]
            try {
                cont := FileRead(info.ruta)
            } catch as e {
                MsgBox(
                    "No se pudo leer el archivo:" . "`n" info.ruta
                    . "`nSe aborta la combinación."
                    . "`nDetalle: " e.Message
                , "Error")
                return
            }

            if (cont = "")
                continue

            cont := NormalizarSaltos(cont)

            if !primeraVez {
                contenidoNuevo .= delim . "; ===========================================" . delim
            } else {
                primeraVez := false
            }

            contenidoNuevo .= "; Archivo origen: " info.nombre . delim . cont
        }

        if (contenidoNuevo = "") {
            MsgBox("No se pudo generar contenido para el archivo combinado (¿todos estaban vacíos?).", "Error")
            return
        }

        ; Proponer nombre por defecto DDMMAAAA-HHMM-script.ini
        ts := FormatTime(, "ddMMyyyy-HHmm")
        nombrePorDefecto := ts "-script.ini"

        ; En AHK v2: InputBox(Prompt, Title?, Options?, Default?)
        ; Dejamos Options vacío (,) y usamos nombrePorDefecto como Default
        res := InputBox(
            "Nombre para el script nuevo (se guardará en 'scriptsCreados'):" . "`n"
          . "Si vacías, se usará el nombre sugerido.",
            "Guardar script nuevo",
            ,  ; Options vacío
            nombrePorDefecto
        )

        if (res.Result = "Cancel") {
            return
        }

        nombreFinal := Trim(res.Value)
        if (nombreFinal = "") {
            nombreFinal := nombrePorDefecto
        }

        ; Asegurar extensión .ini
        if !RegExMatch(nombreFinal, "i)\.ini$") {
            nombreFinal .= ".ini"
        }

        rutaNuevo := scriptsDir "\" nombreFinal

        ; Guardar script nuevo
        try {
            f := FileOpen(rutaNuevo, "w", "UTF-8-RAW")
            if (!f)
                throw Error("No se pudo crear el archivo combinado.")
            f.Write(contenidoNuevo)
            f.Close()
        } catch as e {
            MsgBox("No se pudo guardar el archivo combinado." . "`nDetalle: " e.Message, "Error")
            return
        }

        MsgBox("Se generó el nuevo script:" . "`n" rutaNuevo, "Script creado")

        TransGui.Destroy()
        ProcesarArchivoFinal(rutaNuevo)
    }

    ; --- Función local: seleccionar script ya creado ---
    AbrirScriptsCreados(*) {
        ; Listar scripts en scriptsDir
        listaScripts := []
        Loop Files, scriptsDir "\*.ini" {
            listaScripts.Push({nombre: A_LoopFileName, ruta: A_LoopFileFullPath})
        }

        if (listaScripts.Length = 0) {
            MsgBox("No hay scripts creados en:" . "`n" scriptsDir, "Aviso")
            return
        }

        PrevGui := Gui()
        PrevGui.Title := "Scripts creados"
        PrevGui.SetFont("s10", "Segoe UI")

        PrevGui.Add("Text",, "Seleccioná el script que querés ejecutar:")

        LV_Scripts := PrevGui.Add("ListView", "r15 w520 Grid", ["Archivo"])
        for _, info in listaScripts {
            LV_Scripts.Add("", info.nombre)
        }

        btnOK2     := PrevGui.Add("Button", "y+10 w120 Default", "Usar script")
        btnCancel2 := PrevGui.Add("Button", "x+10 w80", "Cancelar")

        btnOK2.OnEvent("Click", UsarScriptSeleccionado)
        btnCancel2.OnEvent("Click", (*) => PrevGui.Destroy())
        LV_Scripts.OnEvent("DoubleClick", UsarScriptSeleccionado)

        PrevGui.Show()

        UsarScriptSeleccionado(*) {
            fila := LV_Scripts.GetNext()
            if (fila = 0)
                return

            info := listaScripts[fila]
            PrevGui.Destroy()
            TransGui.Destroy()
            ProcesarArchivoFinal(info.ruta)
        }
    }
}

; --------------------------------------------------------------
; PROCESAR ARCHIVO FINAL (SELECCIONADO O COMBINADO)
;   - Guarda ruta en RutaIniSeleccionado
;   - Lee GROUP y CARD
;   - Pregunta si se quieren cambiar
;   - Sí -> IniciarProceso()
;   - No -> CargarYReproducirScript()
; --------------------------------------------------------------
ProcesarArchivoFinal(rutaArchivo) {
    global RutaIniSeleccionado

    RutaIniSeleccionado := rutaArchivo

    try {
        contenido := FileRead(rutaArchivo)
    } catch as e {
        MsgBox("No se pudo leer el archivo seleccionado." . "`nDetalle: " e.Message, "Error")
        return
    }

    if (contenido = "") {
        MsgBox("El archivo seleccionado está vacío.", "Aviso")
        ; Igual seguimos si querés usarlo tal cual
    }

    ; Buscar primeros GROUP= y CARD=
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
    texto .= "**********************************************************************************" . "`n"
    texto .= "GROUP actual: " . (groupVal != "" ? groupVal : "(no encontrado)") . "`n"
    texto .= "CARD  actual: " . (cardVal  != "" ? cardVal  : "(no encontrado)") . "`n`n"
    texto .= "**********************************************************************************" . "`n"
    texto .= "¿Deseás CAMBIAR estos valores según el banco y la tarjeta que selecciones?" . "`n"
    texto .= "(Si elegís 'No', se usará el archivo tal como está y se irá directo a reproducir el script.)"

    resp := MsgBox(texto, "Confirmar acción", "YesNo")

    if (resp = "No") {
        CargarYReproducirScript()
        return
    }

    ; Si respondió Sí, seguimos con el flujo normal de cambio
    IniciarProceso()
}

; --------------------------------------------------------------
; FUNCIÓN 1: INICIA EL PROCESO DE CAMBIO (BANCO/TARJETA)
; --------------------------------------------------------------
IniciarProceso() {
    global TituloVentana, DesplegableBanco_ClassNN
    if !WinExist(TituloVentana) {
        MsgBox("La ventana '" TituloVentana "' no fue encontrada.", "Error")
        return
    }

    try {
        ArrayBancos := ControlGetItems(DesplegableBanco_ClassNN, TituloVentana)
    } catch as e {
        MsgBox("Error al leer la lista de bancos. Verificá el ClassNN." . "`nDetalle: " e.Message, "Error")
        return
    }

    if !ArrayBancos.Length {
        MsgBox("No se encontraron bancos en el desplegable.", "Aviso")
        return
    }

    ElegirBancoGUI(ArrayBancos)
}

; --------------------------------------------------------------
; FUNCIÓN 2: GUI PARA ELEGIR UN BANCO
; --------------------------------------------------------------
ElegirBancoGUI(ArrayBancos) {
    BancoGui := Gui()
    BancoGui.Title := "Paso 1: Seleccioná un Banco"
    BancoGui.SetFont("s10", "Segoe UI")
    BancoGui.Add("Text",, "Seleccioná el banco de la lista:")

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

; --------------------------------------------------------------
; FUNCIÓN 3: SELECCIONA EL BANCO EN LA APP Y LEE LAS TARJETAS
; --------------------------------------------------------------
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
        MsgBox("Error al leer la lista de tarjetas. Verificá el ClassNN." . "`nDetalle: " e.Message, "Error")
        return
    }

    if !ArrayTarjetas.Length {
        MsgBox("No se encontraron tarjetas para este banco.", "Aviso")
        return
    }

    ElegirTarjetaGUI(BancoElegido, ArrayTarjetas)
}

; --------------------------------------------------------------
; FUNCIÓN 4: GUI PARA ELEGIR TARJETA Y MODIFICAR EL INI
; --------------------------------------------------------------
ElegirTarjetaGUI(BancoElegido, ArrayTarjetas) {
    TarjetaGui := Gui()
    TarjetaGui.Title := "Paso 2: Seleccioná una Tarjeta"
    TarjetaGui.SetFont("s10", "Segoe UI")
    TarjetaGui.Add("Text",, "Seleccioná la tarjeta para el banco: " BancoElegido)

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

; --------------------------------------------------------------
; FUNCIÓN 5: MODIFICA GROUP= Y CARD= EN EL ARCHIVO SELECCIONADO
; --------------------------------------------------------------
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

; --------------------------------------------------------------
; FUNCIÓN 6: SCRIPTS -> PLAYBACK SCRIPT FROM -> THIS MACHINE...
; --------------------------------------------------------------
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

        ; b) Fallback: ALT+P
        Send("!p")
    } catch as e {
        MsgBox("No se pudo activar el botón Play." . "`nDetalle: " e.Message, "Error")
    }
}
