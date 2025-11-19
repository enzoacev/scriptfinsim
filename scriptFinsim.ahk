; === GESTOR DE SCRIPTS FINSIM - AHK V2 ===
#Requires AutoHotkey v2.0
#SingleInstance
#Warn

; --- CONFIGURACIÓN GENERAL ---
SetTitleMatchMode 1   ; 1 = el título EMPIEZA con el texto dado
Global TituloVentana := "ATM Simulator"
Global DesplegableBanco_ClassNN := "ComboBox1"
Global ListaTarjetas_ClassNN   := "ListBox1"
Global RutaIniSeleccionado := ""   ; ruta del .ini que se va a ejecutar

; --- ATAJO PRINCIPAL: Ctrl + Alt + T ---
^!t::{
    SeleccionarArchivoYFlujo()
}

; ============================================================
; FUNCIÓN AUXILIAR: normalizar saltos de línea a CRLF (`r`n)
; ============================================================
NormalizarSaltos(texto) {
    if (texto = "")
        return ""
    tmp := StrReplace(texto, "`r`n", "`n")
    tmp := StrReplace(tmp, "`r", "`n")
    return StrReplace(tmp, "`n", "`r`n")
}

; ============================================================
; PANTALLA PRINCIPAL
;   - Lista transacciones base (CODIGO - NombreDeTransaccion)
;   - Permite seleccionar varias con checkboxes
;   - Botón para crear script nuevo combinado
;   - Botón seleccionar/deseleccionar todos
;   - Botón para usar scripts ya generados (scriptsCreados)
; ============================================================
SeleccionarArchivoYFlujo() {
    global RutaIniSeleccionado

    baseDir    := A_ScriptDir
    transDir   := baseDir "\transacciones"
    scriptsDir := transDir "\scriptsCreados"

    if !DirExist(transDir) {
        MsgBox(
            "No se encontró la carpeta 'transacciones'." . "`n`n"
          . "Revisá que la estructura sea:" . "`n"
          . baseDir . "\transacciones",
          "Carpeta no encontrada"
        )
        return
    }

    if !DirExist(scriptsDir) {
        DirCreate(scriptsDir)
    }

    ; Mover cualquier script generado viejo que esté suelto a scriptsCreados
    Loop Files, transDir "\*.ini" {
        nombre := A_LoopFileName
        if RegExMatch(nombre, "^\d{8}-\d{4}-script\.ini$") {
            FileMove(A_LoopFileFullPath, scriptsDir "\" nombre, 1)
        }
    }

    ; --- Cargar SOLO transacciones base (formato: CODIGO - Nombre...) ---
    ;   Ej: 101100 - ExtraccionCA-Pesos.ini
    archivos := []  ; {nombre, ruta}
    Loop Files, transDir "\*.ini" {
        nombre := A_LoopFileName
        ruta   := A_LoopFileFullPath

        ; CODIGO - Nombre
        if RegExMatch(nombre, "i)^\s*\d+\s*-\s+(.+)$") {
            archivos.Push({nombre: nombre, ruta: ruta})
        }
    }

    if (archivos.Length = 0) {
        MsgBox(
            "No se encontraron transacciones base en la carpeta 'transacciones'." . "`n`n"
          . "Recordá que el formato debe ser:" . "`n"
          . "CODIGO - NombreDeTransaccion.ini" . "`n"
          . "Ejemplo: 101100 - ExtraccionCA-Pesos.ini",
          "Sin transacciones base"
        )
        return
    }

    ; Orden alfabético por nombre de archivo
    for i, _ in archivos {
        if (i >= archivos.Length)
            break
        j := i + 1
        while (j <= archivos.Length) {
            a := archivos[i]
            b := archivos[j]
            if (StrCompare(a.nombre, b.nombre) > 0) {
                temp := archivos[i]
                archivos[i] := archivos[j]
                archivos[j] := temp
            }
            j++
        }
    }

    ; --- GUI PRINCIPAL ---
    TransGui := Gui("+AlwaysOnTop")
    TransGui.Opt("+OwnDialogs")          ; <<< ESTA LÍNEA NUEVA
    TransGui.Title := "Gestor de scripts FinSim"

    ; Fuente normal base
    TransGui.SetFont("s10", "Segoe UI")


    ; Título en negrita
    TransGui.SetFont("s10 Bold", "Segoe UI")
    TransGui.Add("Text", "w520 Center", "Selección de transacciones base")

    ; Descripción en normal
    TransGui.SetFont("s9", "Segoe UI")
    TransGui.Add("Text", "w520", "Marcá una o varias transacciones (CODIGO - NombreDeTransaccion) para generar un nuevo script:")

    LV_Arch := TransGui.Add("ListView", "r15 w520 Checked Grid", ["Archivo"])
    for idx, info in archivos {
        LV_Arch.Add("", info.nombre)
    }

    ; Botones
    btnCombinar     := TransGui.Add("Button", "y+10 w160 Default", "Crear script nuevo")
    btnSelectAll    := TransGui.Add("Button", "x+10 w190", "Seleccionar / deseleccionar todo")
    btnScriptPrevio := TransGui.Add("Button", "x+10 w150", "Ejecutar script guardado")
    btnCancel       := TransGui.Add("Button", "x+10 w90", "Cerrar")

    btnCombinar.OnEvent("Click", ConfirmarSeleccionBase)
    btnSelectAll.OnEvent("Click", SeleccionarTodo)
    btnScriptPrevio.OnEvent("Click", AbrirScriptsCreados)
    btnCancel.OnEvent("Click", (*) => TransGui.Destroy())
    LV_Arch.OnEvent("DoubleClick", ToggleCheckbox)

    TransGui.Show()

    ; -------- Funciones locales de la GUI principal --------

    ; Toggle de check con doble clic
    ToggleCheckbox(*) {
        fila := LV_Arch.GetNext()
        if (fila = 0)
            return
        estado := LV_Arch.GetChecked(fila)
        LV_Arch.Modify(fila, estado ? "-Check" : "Check")
    }

    ; Seleccionar / deseleccionar todas las filas
    SeleccionarTodo(*) {
        total := LV_Arch.GetCount()
        if (total = 0)
            return

        checked := 0
        row := 0
        while (row := LV_Arch.GetNext(row, "C")) { ; "C" = Checked
            checked++
        }

        if (checked = total) {
            ; Todas marcadas -> desmarcar todas
            Loop total {
                LV_Arch.Modify(A_Index, "-Check")
            }
        } else {
            ; Faltan algunas -> marcar todas
            Loop total {
                LV_Arch.Modify(A_Index, "Check")
            }
        }
    }

    ; Crear script nuevo combinado a partir de las transacciones marcadas
        ; Crear script nuevo combinado a partir de las transacciones marcadas
    ConfirmarSeleccionBase(*) {
        seleccionados := []
        row := 0

        while (row := LV_Arch.GetNext(row, "C")) {
            seleccionados.Push(row)
        }

        if (seleccionados.Length = 0) {
            MsgBox(
                "No marcaste ninguna transacción base." . "`n`n"
              . "Marcá al menos una para poder generar un script.",
              "Nada seleccionado"
            )
            return
        }

        delim           := "`r`n"
        contenidoNuevo  := ""
        primeraVez      := true

        for _, rowIdx in seleccionados {
            info := archivos[rowIdx]
            try {
                cont := FileRead(info.ruta)
            } catch as e {
                MsgBox(
                    "No se pudo leer el archivo:" . "`n" info.ruta . "`n`n"
                  . "Se cancela la generación del script." . "`n"
                  . "Detalle técnico: " e.Message,
                  "Error al leer archivo"
                )
                return
            }

            if (cont = "")
                continue

            cont := NormalizarSaltos(cont)

            if !primeraVez {
                contenidoNuevo .= delim . "; ================= TRANSACCIÓN SIGUIENTE =================" . delim
            } else {
                primeraVez := false
            }

            contenidoNuevo .= "; Origen: " info.nombre . delim . cont
        }

        if (contenidoNuevo = "") {
            MsgBox(
                "El contenido combinado resultó vacío." . "`n`n"
              . "Revisá que las transacciones base no estén vacías.",
              "Contenido vacío"
            )
            return
        }

        ; Sugerir nombre DDMMAAAA-HHMM-script.ini
        ts := FormatTime(, "ddMMyyyy-HHmm")
        nombrePorDefecto := ts "-script.ini"

        ; --- OCULTAR LA GUI PRINCIPAL MIENTRAS SE MUESTRA EL INPUTBOX ---
        TransGui.Hide()

        res := InputBox(
            "Ingresá el nombre para el nuevo script (se guardará en 'scriptsCreados'):" . "`n`n"
          . "Si dejás el campo vacío, se usará el nombre sugerido:" . "`n"
          . nombrePorDefecto,
            "Guardar script nuevo",
            ,                ; Options vacío
            nombrePorDefecto ; Default
        )

        ; Si canceló, volvemos a mostrar la ventana y salimos
        if (res.Result = "Cancel") {
            TransGui.Show()
            WinActivate(TransGui)   ; por si quedó detrás de otra ventana
            return
        }

        nombreFinal := Trim(res.Value)
        if (nombreFinal = "") {
            nombreFinal := nombrePorDefecto
        }

        if !RegExMatch(nombreFinal, "i)\.ini$") {
            nombreFinal .= ".ini"
        }

        rutaNuevo := scriptsDir "\" nombreFinal

        try {
            f := FileOpen(rutaNuevo, "w", "UTF-8-RAW")
            if (!f)
                throw Error("No se pudo crear el archivo en disco.")
            f.Write(contenidoNuevo)
            f.Close()
        } catch as e {
            ; En caso de error, volvemos a mostrar la GUI principal
            TransGui.Show()
            WinActivate(TransGui)

            MsgBox(
                "No se pudo guardar el nuevo script en:" . "`n" rutaNuevo . "`n`n"
              . "Detalle técnico: " e.Message,
              "Error al guardar script"
            )
            return
        }

        MsgBox(
            "El script se generó correctamente:" . "`n" rutaNuevo,
            "Script creado"
        )

        ; Ya no necesitamos la GUI principal
        TransGui.Destroy()
        ProcesarArchivoFinal(rutaNuevo)
    }


    ; Listar y elegir scripts previamente generados
    AbrirScriptsCreados(*) {
        listaScripts := []
        Loop Files, scriptsDir "\*.ini" {
            listaScripts.Push({nombre: A_LoopFileName, ruta: A_LoopFileFullPath})
        }

        if (listaScripts.Length = 0) {
            MsgBox(
                "No se encontraron scripts guardados en:" . "`n" scriptsDir,
                "Sin scripts guardados"
            )
            return
        }

        PrevGui := Gui("+AlwaysOnTop")
        PrevGui.Title := "Scripts guardados"

        PrevGui.SetFont("s10", "Segoe UI")
        PrevGui.SetFont("s10 Bold", "Segoe UI")
        PrevGui.Add("Text", "w520 Center", "Seleccioná un script guardado para ejecutar")
        PrevGui.SetFont("s9", "Segoe UI")

        LV_Scripts := PrevGui.Add("ListView", "r15 w520 Grid", ["Archivo"])
        for _, info in listaScripts {
            LV_Scripts.Add("", info.nombre)
        }

        btnOK2     := PrevGui.Add("Button", "y+10 w150 Default", "Ejecutar script")
        btnCancel2 := PrevGui.Add("Button", "x+10 w90", "Cancelar")

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

; ============================================================
; PROCESAR ARCHIVO SELECCIONADO (nuevo o guardado)
;   - Muestra resumen de GROUP/CARD
;   - Pregunta si querés actualizarlos según banco/tarjeta
; ============================================================
ProcesarArchivoFinal(rutaArchivo) {
    global RutaIniSeleccionado

    RutaIniSeleccionado := rutaArchivo

    try {
        contenido := FileRead(rutaArchivo)
    } catch as e {
        MsgBox(
            "No se pudo leer el archivo seleccionado:" . "`n" rutaArchivo . "`n`n"
          . "Detalle técnico: " e.Message,
          "Error al leer archivo"
        )
        return
    }

    if (contenido = "") {
        MsgBox(
            "El archivo seleccionado está vacío:" . "`n" rutaArchivo,
            "Archivo vacío"
        )
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
    texto .= "Valores detectados:" . "`n"
    texto .= "  GROUP = " . (groupVal != "" ? groupVal : "(no encontrado)") . "`n"
    texto .= "  CARD  = " . (cardVal  != "" ? cardVal  : "(no encontrado)") . "`n`n"
    texto .= "¿Querés ACTUALIZAR estos valores según el banco y la tarjeta que selecciones en el ATM Simulator?" . "`n`n"
    texto .= "Sí  → Se abre el ATM, elegís banco y tarjeta, y se actualizan GROUP/CARD en el archivo." . "`n"
    texto .= "No  → Se ejecuta el script tal como está."

    resp := MsgBox(texto, "¿Actualizar GROUP/CARD?", "YesNo")

    if (resp = "No") {
        CargarYReproducirScript()
        return
    }

    IniciarProceso()
}

; ============================================================
; PASO 1: leer bancos desde el ATM y mostrarlos en GUI
; ============================================================
IniciarProceso() {
    global TituloVentana, DesplegableBanco_ClassNN

    if !WinExist(TituloVentana) {
        MsgBox(
            "No se encontró la ventana del ATM Simulator." . "`n`n"
          . "Asegurate de que esté abierta y que el título comience con:" . "`n"
          . "'" TituloVentana "'",
          "ATM Simulator no encontrado"
        )
        return
    }

    try {
        ArrayBancos := ControlGetItems(DesplegableBanco_ClassNN, TituloVentana)
    } catch as e {
        MsgBox(
            "Ocurrió un error al leer la lista de bancos del ATM." . "`n`n"
          . "Verificá el ClassNN del combo de bancos." . "`n"
          . "Detalle técnico: " e.Message,
          "Error leyendo bancos"
        )
        return
    }

    if !ArrayBancos.Length {
        MsgBox(
            "No se encontraron bancos en el desplegable del ATM.",
            "Lista de bancos vacía"
        )
        return
    }

    ElegirBancoGUI(ArrayBancos)
}

; ============================================================
; GUI: selección de banco
; ============================================================
ElegirBancoGUI(ArrayBancos) {
    BancoGui := Gui("+AlwaysOnTop")
    BancoGui.Title := "Paso 1 - Selección de banco"

    BancoGui.SetFont("s10", "Segoe UI")
    BancoGui.SetFont("s10 Bold", "Segoe UI")
    BancoGui.Add("Text",, "Elegí el banco que corresponde al script")
    BancoGui.SetFont("s9", "Segoe UI")
    BancoGui.Add("Text",, "Hacé doble clic o presioná Aceptar con un banco seleccionado:")

    LV_Bancos := BancoGui.Add("ListView", "r15 w420 Sort", ["Bancos disponibles"])
    LV_Bancos.OnEvent("DoubleClick", BancoSeleccionado)

    for _, banco in ArrayBancos {
        LV_Bancos.Add("", Trim(banco))
    }

    btnOK := BancoGui.Add("Button", "y+10 w100 Default", "Aceptar")
    btnOK.OnEvent("Click", BancoSeleccionado)

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

; ============================================================
; PASO 2: seleccionar banco en ATM, leer tarjetas y mostrar GUI
; ============================================================
ContinuarConBanco(BancoElegido) {
    global TituloVentana, DesplegableBanco_ClassNN, ListaTarjetas_ClassNN

    ; Activar ventana del ATM
    WinActivate(TituloVentana)
    WinWaitActive(TituloVentana, , 2)

    ; Seleccionar el banco en el combo del ATM
    try {
        ControlChooseString(BancoElegido, DesplegableBanco_ClassNN, TituloVentana)
    } catch as e {
        MsgBox(
            "No se pudo seleccionar el banco '" BancoElegido "' en el ATM." . "`n`n"
          . "Detalle técnico: " e.Message,
          "Error seleccionando banco"
        )
        return
    }

    ; === MOSTRAR MENSAJE DE ESPERA MIENTRAS TRAE LAS TARJETAS ===
    WaitGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    WaitGui.Opt("+OwnDialogs")
    ; márgenes de la ventana
    WaitGui.MarginX := 10
    WaitGui.MarginY := 10

    WaitGui.SetFont("s10", "Segoe UI")
    ; Solo usamos "Center" como opción del Text
    WaitGui.Add(
        "Text",
        "Center",
        "Aguarda un momento..." . "`n"
      . "Estamos recuperando las tarjetas del banco seleccionado."
    )
    WaitGui.Show("AutoSize Center")

    try {
        ; Pequeña pausa para que el ATM cargue las tarjetas
        Sleep(2000)

        ; Leer tarjetas desde la lista del ATM
        try {
            ArrayTarjetas := ControlGetItems(ListaTarjetas_ClassNN, TituloVentana)
        } catch as eInner {
            MsgBox(
                "Ocurrió un error al leer la lista de tarjetas del ATM." . "`n`n"
              . "Verificá el ClassNN del listado de tarjetas." . "`n"
              . "Detalle técnico: " eInner.Message,
              "Error leyendo tarjetas"
            )
            return
        }

        if !ArrayTarjetas.Length {
            MsgBox(
                "No se encontraron tarjetas para el banco seleccionado.",
                "Lista de tarjetas vacía"
            )
            return
        }

        ; Si todo salió bien, pasamos a la GUI de tarjetas
        ElegirTarjetaGUI(BancoElegido, ArrayTarjetas)

    } finally {
        ; Siempre cerrar la ventanita de espera, haya salido bien o mal
        try WaitGui.Destroy()
    }
}


; ============================================================
; GUI: selección de tarjeta
; ============================================================
ElegirTarjetaGUI(BancoElegido, ArrayTarjetas) {
    TarjetaGui := Gui("+AlwaysOnTop")
    TarjetaGui.Title := "Paso 2 - Selección de tarjeta"

    TarjetaGui.SetFont("s10", "Segoe UI")
    TarjetaGui.SetFont("s10 Bold", "Segoe UI")
    TarjetaGui.Add("Text",, "Banco seleccionado: " BancoElegido)
    TarjetaGui.SetFont("s9", "Segoe UI")
    TarjetaGui.Add("Text",, "Elegí la tarjeta que querés usar en el script:")

    LV_Tarjetas := TarjetaGui.Add("ListView", "r15 w520", ["Tarjetas disponibles"])
    LV_Tarjetas.OnEvent("DoubleClick", TarjetaSeleccionada)

    for _, tarjeta in ArrayTarjetas {
        LV_Tarjetas.Add("", Trim(tarjeta))
    }

    btnOK := TarjetaGui.Add("Button", "y+10 w120 Default", "Usar tarjeta")
    btnOK.OnEvent("Click", TarjetaSeleccionada)

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

; ============================================================
; ACTUALIZAR GROUP= y CARD= en el INI
; ============================================================
ModificarArchivoIni(Grupo, Tarjeta) {
    global RutaIniSeleccionado

    rutaArchivo := RutaIniSeleccionado
    if (rutaArchivo = "") {
        MsgBox("No hay ningún archivo .ini seleccionado para modificar.", "Sin archivo")
        return
    }

    try {
        original := FileRead(rutaArchivo)
        if (original = "")
            throw Error("El archivo está vacío o no se pudo leer.")
    } catch as e {
        MsgBox(
            "No se pudo leer el archivo para modificar GROUP/CARD:" . "`n" rutaArchivo . "`n`n"
          . "Detalle técnico: " e.Message,
          "Error al leer archivo"
        )
        return
    }

    delim    := InStr(original, "`r`n") ? "`r`n" : "`n"
    trailing := (StrLen(original) >= StrLen(delim)) && (SubStr(original, -StrLen(delim) + 1) = delim)

    lines := StrSplit(original, delim)
    totalGroup := 0
    totalCard  := 0

    for i, l in lines {
        cG := 0
        l := RegExReplace(
            l
          , "i)^( *|\t*)(GROUP)(\s*)=.*$"
          , "$1$2$3=" Grupo
          , &cG
        )
        if (cG)
            totalGroup += cG

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
        MsgBox(
            "No se encontraron claves 'GROUP=' ni 'CARD=' en el archivo." . "`n`n"
          . "No se realizaron cambios.",
          "Sin coincidencias"
        )
    } else {
        out := ""
        for i, l in lines {
            if (i > 1)
                out .= delim
            out .= l
        }
        if (trailing)
            out .= delim

        try {
            FileCopy(rutaArchivo, rutaArchivo ".bak", true)
            f := FileOpen(rutaArchivo, "w", "UTF-8-RAW")
            if (!f)
                throw Error("No se pudo abrir el archivo para escritura.")
            f.Write(out)
            f.Close()
        } catch as e {
            MsgBox(
                "No se pudo escribir el archivo actualizado." . "`n`n"
              . "Detalle técnico: " e.Message,
              "Error al guardar cambios"
            )
            return
        }

        MsgBox(
            "Actualización completada:" . "`n`n"
          . "GROUP = " Grupo "  (reemplazos: " totalGroup ")" . "`n"
          . "CARD  = " Tarjeta "  (reemplazos: " totalCard ")",
          "GROUP/CARD actualizados"
        )
    }

    CargarYReproducirScript()
}

; ============================================================
; CARGAR Y EJECUTAR EL SCRIPT EN EL ATM SIMULATOR
; ============================================================
CargarYReproducirScript() {
    global TituloVentana, RutaIniSeleccionado

    if (RutaIniSeleccionado = "") {
        MsgBox(
            "No hay ruta de archivo .ini para ejecutar.",
            "Sin archivo seleccionado"
        )
        return
    }

    WinActivate(TituloVentana)
    if !WinWaitActive(TituloVentana, , 2) {
        MsgBox(
            "No se pudo activar la ventana del ATM Simulator." . "`n`n"
          . "Verificá que esté abierta y visible.",
          "ATM Simulator no activo"
        )
        return
    }

    ; Menú: Scripts -> Playback Script From -> This machine...
    Send("!s")
    Sleep(250)
    Send("s")
    Sleep(250)
    Send("t")

    if !WinWaitActive("ahk_class #32770", , 3) {
        MsgBox(
            "No se detectó la ventana estándar de selección de archivo (Abrir/Open).",
            "Diálogo no encontrado"
        )
        return
    }

    try {
        path := RutaIniSeleccionado
        path := Trim(path, '"')

        ControlFocus("Edit1", "ahk_class #32770")
        Sleep(150)
        Send("^a")
        Sleep(80)
        Send("{Del}")
        Sleep(80)
        SendText(path)
        Sleep(200)
        Send("{Enter}")
    } catch as e {
        MsgBox(
            "Ocurrió un error al completar la ruta del archivo en el diálogo de apertura." . "`n`n"
          . "Detalle técnico: " e.Message,
          "Error al seleccionar archivo"
        )
        return
    }

    playWin := "Play Back A Local Script File"
    if !WinWait(playWin, , 5) {
        MsgBox(
            "No se encontró la ventana 'Play Back A Local Script File' para iniciar la reproducción.",
            "Ventana de Play no encontrada"
        )
        return
    }

    WinActivate(playWin)
    WinWaitActive(playWin, , 2)

    try {
        ControlFocus("Button12", playWin)
        Sleep(150)
        Send("{Space}")
        Sleep(300)
        Send("!p")  ; Fallback por si tiene Alt+P
    } catch as e {
        MsgBox(
            "No se pudo activar el botón Play en la ventana de reproducción." . "`n`n"
          . "Detalle técnico: " e.Message,
          "Error al ejecutar script"
        )
    }
}
