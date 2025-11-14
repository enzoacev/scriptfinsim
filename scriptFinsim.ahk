; === SCRIPT DEFINITIVO Y CORREGIDO - AHK V2 ===
#Requires AutoHotkey v2.0
#SingleInstance
#Warn

; --- CONFIGURACIÓN ---
; ¡ASEGÚRATE DE QUE ESTOS 3 VALORES SEAN EXACTOS!
SetTitleMatchMode 1   ; 1 = el título EMPIEZA con el texto dado
Global TituloVentana := "ATM Simulator"
Global DesplegableBanco_ClassNN := "ComboBox1"
Global ListaTarjetas_ClassNN   := "ListBox1"

; --- Atajo de teclado para iniciar: Presiona Ctrl + Alt + T ---
^!t::{
    IniciarProceso()
}

; --- FUNCIÓN 1: INICIA TODO EL PROCESO ---
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

; --- FUNCIÓN 2: MUESTRA LA GUI PARA ELEGIR UN BANCO ---
ElegirBancoGUI(ArrayBancos) {
    BancoGui := Gui()
    BancoGui.Title := "Paso 1: Selecciona un Banco"
    BancoGui.SetFont("s10", "Segoe UI")
    BancoGui.Add("Text",, "Selecciona el banco de la lista:")

    LV_Bancos := BancoGui.Add("ListView", "r15 w400 Sort", ["Bancos Disponibles"])
    LV_Bancos.OnEvent("DoubleClick", BancoSeleccionado)

    ; Cargar bancos (limpiando espacios extremos)
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

; --- FUNCIÓN 4: MUESTRA LA GUI PARA ELEGIR UNA TARJETA Y MODIFICA EL INI ---
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

; --- FUNCIÓN 5: cambia TODAS las apariciones de GROUP= y CARD= en el archivo ---
ModificarArchivoIni(Grupo, Tarjeta) {
    rutaArchivo := FileSelect(, , "Selecciona el archivo .ini a modificar", "Archivos INI (*.ini)")
    if (rutaArchivo = "")
        return

    try {
        original := FileRead(rutaArchivo)
        if (original = "")
            throw Error("El archivo está vacío o no se pudo leer.")

        ; Detectar separador original y si el archivo terminaba con salto
        delim    := InStr(original, "`r`n") ? "`r`n" : "`n"
        trailing := (StrLen(original) >= StrLen(delim)) && (SubStr(original, -StrLen(delim) + 1) = delim)

        lines := StrSplit(original, delim)

        totalGroup := 0, totalCard := 0

        ; Reemplaza en TODAS las líneas (evita tocar comentarios porque no matchean ^\s*KEY=)
        for i, l in lines {
            ; GROUP=
            cG := 0
            ; Conserva indentación ($1), la clave tal como esté ($2) y espacios antes del '=' ($3)
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
            return
        }

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
            "Hecho: se actualizaron todas las apariciones." 
            . "`nGROUP = " Grupo "  (reemplazos: " totalGroup ")"
            . "`nCARD  = " Tarjeta "  (reemplazos: " totalCard ")"
        , "OK")

        ; === NUEVO: luego de modificar el archivo, cargarlo y darle Play en el simulador ===
        CargarYReproducirScript()

    } catch as e {
        MsgBox("No se pudo modificar el archivo." . "`nDetalle: " e.Message, "Error")
    }
}

; --- FUNCIÓN 6: va al simulador, Scripts -> Playback Script From -> This machine... y da Play ---
CargarYReproducirScript() {
    global TituloVentana

    ; Aseguramos foco en el ATM Simulator
    WinActivate(TituloVentana)
    if !WinWaitActive(TituloVentana, , 2) {
        MsgBox("No se pudo activar la ventana del ATM Simulator para cargar el script.", "Error")
        return
    }

    ; 1) Abrir menú Scripts (ALT + S)
    Send("!s")
    Sleep(250)

    ; 2) Elegir 'Playback Script From' (letra aceleradora S de Script)
    Send("s")
    Sleep(250)

    ; 3) En el submenú elegir 'This machine...' (letra aceleradora T)
    Send("t")

    ; Ahora se abre el Explorador para que vos elijas el archivo
    MsgBox(
        "Se abrió la ventana del Explorador para seleccionar el script." 
        . "`nSeleccioná el archivo correspondiente y presioná Aceptar." 
        . "`n`nCuando la ventana se cierre y vuelvas al simulador, presioná OK para continuar y darle Play.",
        "Seleccionar script"
    )

    ; 4) Volver al ATM Simulator y dar Play en 'Play Back A Local Script File'
    WinActivate(TituloVentana)
    if !WinWaitActive(TituloVentana, , 3) {
        MsgBox("No se pudo volver al ATM Simulator después de seleccionar el archivo.", "Error")
        return
    }

    ; 5) Hacer clic en el botón Play (Button12)
    try {
        ControlClick("Button12", TituloVentana)
    } catch as e {
        MsgBox("No se pudo hacer clic en el botón Play (Button12)." . "`nDetalle: " e.Message, "Error")
    }
}
