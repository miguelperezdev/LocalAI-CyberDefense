Option Explicit

' ============================================================
' PROYECTO: CiberDefensa con IA Local (Ollama) - Versión Estable
' AUTOR: Miguel Perez
' DESCRIPCIÓN: Analiza eventos con Ollama si está disponible,
'              sino usa heurística. Con barra de progreso y cancelación.
' ============================================================

' ---------- CONFIGURACIÓN ----------
Private Const OLLAMA_URL As String = "http://localhost:11434/api/generate"
Private Const OLLAMA_MODEL As String = "llama3"   ' Cambia si usas otro (mistral, etc.)
Private Const TIMEOUT_SEGUNDOS As Long = 8        ' Timeout para cada llamada

' ---------- VARIABLE GLOBAL PARA CANCELAR ----------
Private CancelarProceso As Boolean

' ============================================================
' FUNCIÓN PRINCIPAL (LA QUE EJECUTAS)
' ============================================================
Sub EjecutarAnalisisIA()
    Dim wsMon As Worksheet
    Dim wsDash As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim eventoTexto As String
    Dim riesgo As String, accion As String, motivo As String
    Dim recomendaciones As String
    Dim hayOllama As Boolean
    Dim respuestaIA As String

    ' Inicializar
    CancelarProceso = False
    Set wsMon = ThisWorkbook.Sheets("Monitoreo")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    ' Verificar si Ollama está disponible (timeout corto)
    hayOllama = VerificarOllamaRapido()
    If Not hayOllama Then
        If MsgBox("Ollama no responde. ¿Usar solo heurística?", vbYesNo + vbQuestion, "Modo offline") = vbNo Then
            Exit Sub
        End If
    End If

    ' Última fila con datos
    ultimaFila = wsMon.Cells(wsMon.Rows.Count, 1).End(xlUp).Row
    If ultimaFila < 2 Then
        MsgBox "No hay datos en la hoja Monitoreo.", vbExclamation
        Exit Sub
    End If

    ' Limpiar resultados anteriores
    wsMon.Range("F2:I" & ultimaFila).ClearContents
    wsMon.Range("F2:I" & ultimaFila).Interior.ColorIndex = xlNone

    Application.ScreenUpdating = False

    ' Bucle principal
    For i = 2 To ultimaFila
        ' Permitir cancelar con Ctrl+Break o desde el botón
        DoEvents
        If CancelarProceso Then
            MsgBox "Proceso cancelado por el usuario.", vbInformation
            Exit For
        End If

        ' Actualizar barra de estado
        Application.StatusBar = "Analizando fila " & i & " de " & ultimaFila & "..."

        ' Construir evento
        eventoTexto = "Usuario=" & wsMon.Cells(i, 2).Value & _
                      ", IP=" & wsMon.Cells(i, 3).Value & _
                      ", Hora=" & Format(wsMon.Cells(i, 4).Value, "hh:mm") & _
                      ", Fallos=" & wsMon.Cells(i, 5).Value

        ' --- Llamar a IA (si está disponible) ---
        If hayOllama Then
            respuestaIA = LlamarOllama(eventoTexto)
        Else
            respuestaIA = ""
        End If

        ' Si la IA falla o no hay Ollama, usar heurística
        If respuestaIA = "" Then
            respuestaIA = AnalisisHeuristico(i)
        End If

        ' Parsear resultado (formato: "RIESGO|ACCIÓN|MOTIVO|RECOMENDACIONES")
        Dim partes As Variant
        partes = Split(respuestaIA, "|")
        If UBound(partes) >= 2 Then
            riesgo = partes(0)
            accion = partes(1)
            motivo = partes(2)
            If UBound(partes) >= 3 Then
                recomendaciones = partes(3)
            Else
                recomendaciones = ""
            End If
        Else
            riesgo = "ERROR"
            accion = "Revisar"
            motivo = "Respuesta inválida"
            recomendaciones = ""
        End If

        ' Guardar en hoja
        wsMon.Cells(i, 6).Value = motivo                 ' F: motivo
        wsMon.Cells(i, 7).Value = riesgo                 ' G: riesgo
        wsMon.Cells(i, 8).Value = accion                 ' H: acción
        wsMon.Cells(i, 9).Value = recomendaciones        ' I: recomendaciones extra

        ' Colorear fila según riesgo
        Select Case riesgo
            Case "BAJO"
                wsMon.Rows(i).Interior.Color = RGB(200, 255, 200)
            Case "MEDIO"
                wsMon.Rows(i).Interior.Color = RGB(255, 255, 150)
            Case "ALTO"
                wsMon.Rows(i).Interior.Color = RGB(255, 150, 150)
            Case "CRITICO"
                wsMon.Rows(i).Interior.Color = RGB(255, 100, 100)
                ' Registrar alerta
                RegistrarAlerta wsMon.Cells(i, 1).Value, wsMon.Cells(i, 2).Value, wsMon.Cells(i, 3).Value, motivo
        End Select
    Next i

    ' Actualizar dashboard
    ActualizarDashboard ultimaFila - 1

    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Análisis completado.", vbInformation, "Ciberdefensa"
End Sub

' ============================================================
' FUNCIÓN PARA CANCELAR (ASIGNAR A UN BOTÓN)
' ============================================================
Sub CancelarAnalisis()
    CancelarProceso = True
    Application.StatusBar = "Cancelando..."
End Sub

' ============================================================
' VERIFICAR OLLAMA (RÁPIDO)
' ============================================================
Private Function VerificarOllamaRapido() As Boolean
    Dim http As Object
    On Error GoTo ErrorHandler
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 2000, 2000, 2000, 2000   ' 2 segundos máximo
    http.Open "GET", "http://localhost:11434/api/tags", False
    http.Send
    If http.Status = 200 Then
        VerificarOllamaRapido = True
    Else
        VerificarOllamaRapido = False
    End If
    Exit Function
ErrorHandler:
    VerificarOllamaRapido = False
End Function

' ============================================================
' LLAMADA A OLLAMA (CON TIMEOUT CORTO)
' ============================================================
Private Function LlamarOllama(evento As String) As String
    Dim http As Object
    Dim requestBody As String
    Dim prompt As String
    Dim respuesta As String
    Dim inicio As Double

    ' --- Prompt simple (sin JSON complejo) para evitar errores de parseo ---
    prompt = "Eres un analista SOC. Analiza este evento y responde SOLO con una línea en este formato exacto: " & _
             "RIESGO|ACCIÓN|MOTIVO|RECOMENDACIONES. " & _
             "Ejemplo: ALTO|Bloquear IP|Múltiples fallos desde IP sospechosa|Revisar logs,notificar a admin. " & _
             "Evento: " & evento

    ' Escapar comillas para JSON
    prompt = Replace(prompt, "\", "\\")
    prompt = Replace(prompt, """", "\""")

    requestBody = "{""model"":""" & OLLAMA_MODEL & """," & _
                  """prompt"":""" & prompt & """," & _
                  """stream"":false," & _
                  """options"":{""temperature"":0.3}}"

    On Error GoTo ErrorHandler
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' Timeout total: TIMEOUT_SEGUNDOS
    http.SetTimeouts TIMEOUT_SEGUNDOS * 1000, TIMEOUT_SEGUNDOS * 1000, _
                     TIMEOUT_SEGUNDOS * 1000, TIMEOUT_SEGUNDOS * 1000

    inicio = Timer
    http.Open "POST", OLLAMA_URL, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send requestBody

    If Timer - inicio > TIMEOUT_SEGUNDOS Then GoTo ErrorHandler   ' Timeout manual

    If http.Status = 200 Then
        respuesta = http.responseText
        LlamarOllama = ExtraerRespuestaSimple(respuesta)
    Else
        LlamarOllama = ""
    End If
    Exit Function

ErrorHandler:
    LlamarOllama = ""
End Function

' ============================================================
' EXTRAER RESPUESTA DE OLLAMA (SIN PARSER COMPLEJO)
' ============================================================
Private Function ExtraerRespuestaSimple(json As String) As String
    Dim posIni As Long, posFin As Long
    Dim resp As String

    ' Buscar "response":"...
    posIni = InStr(json, """response"":""") + 12
    If posIni = 12 Then
        ExtraerRespuestaSimple = ""
        Exit Function
    End If

    posFin = InStr(posIni, json, """") - 1
    If posFin < posIni Then
        ExtraerRespuestaSimple = ""
        Exit Function
    End If

    resp = Mid(json, posIni, posFin - posIni + 1)
    ' Reemplazar secuencias escapadas
    resp = Replace(resp, "\n", "")
    resp = Replace(resp, "\r", "")
    resp = Replace(resp, "\""", """")

    ' Limpiar y devolver
    ExtraerRespuestaSimple = Trim(resp)
End Function

' ============================================================
' HEURÍSTICA MEJORADA (RÁPIDA)
' ============================================================
Private Function AnalisisHeuristico(fila As Long) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Monitoreo")

    Dim usuario As String, ip As String
    Dim fallos As Long, hora As Long
    Dim score As Long
    Dim riesgo As String, accion As String, motivo As String, recomendaciones As String

    usuario = ws.Cells(fila, 2).Value
    ip = ws.Cells(fila, 3).Value
    fallos = Val(ws.Cells(fila, 5).Value)

    If IsDate(ws.Cells(fila, 4).Value) Then
        hora = Hour(ws.Cells(fila, 4).Value)
    Else
        hora = 12
    End If

    ' Scoring
    score = fallos * 8
    If hora < 7 Or hora > 20 Then score = score + 15
    If ip Like "192.168.1.*" Then score = score + 5
    If ip = "10.0.0.50" Or ip = "203.0.113.45" Then score = score + 40
    If LCase(usuario) Like "*admin*" Or LCase(usuario) = "root" Then score = score + 25

    ' Clasificación
    If score < 20 Then
        riesgo = "BAJO"
        accion = "Monitorizar"
        motivo = "Actividad normal"
        recomendaciones = "Continuar monitoreo"
    ElseIf score < 40 Then
        riesgo = "MEDIO"
        accion = "Revisar logs"
        motivo = "Patrón sospechoso"
        recomendaciones = "Verificar horario; revisar IP"
    ElseIf score < 70 Then
        riesgo = "ALTO"
        accion = "Bloquear temporalmente"
        motivo = "Múltiples fallos y/o IP sospechosa"
        recomendaciones = "Aplicar bloqueo por 1 hora; notificar a admin"
    Else
        riesgo = "CRITICO"
        accion = "Aislar segmento"
        motivo = "Amenaza inminente"
        recomendaciones = "Desconectar equipo; análisis forense"
    End If

    AnalisisHeuristico = riesgo & "|" & accion & "|" & motivo & "|" & recomendaciones
End Function

' ============================================================
' FUNCIONES DE DASHBOARD
' ============================================================

Private Sub RegistrarAlerta(eventoId As String, usuario As String, ip As String, motivo As String)
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(ultimaFila, 1).Value = Now
    ws.Cells(ultimaFila, 2).Value = "?? ALERTA CRÍTICA: " & motivo
    ws.Cells(ultimaFila, 3).Value = usuario
    ws.Cells(ultimaFila, 4).Value = ip
End Sub

Private Sub ActualizarDashboard(totalEventos As Long)
    Dim wsDash As Worksheet
    Dim wsMon As Worksheet
    Dim conteoMedio As Long, conteoAlto As Long, conteoCritico As Long

    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsMon = ThisWorkbook.Sheets("Monitoreo")

    conteoMedio = Application.WorksheetFunction.CountIf(wsMon.Range("G:G"), "MEDIO")
    conteoAlto = Application.WorksheetFunction.CountIf(wsMon.Range("G:G"), "ALTO")
    conteoCritico = Application.WorksheetFunction.CountIf(wsMon.Range("G:G"), "CRITICO")

    ' Actualizar resumen
    wsDash.Range("E1").Value = "RESUMEN DE SEGURIDAD"
    wsDash.Range("E2").Value = "Eventos analizados:"
    wsDash.Range("F2").Value = totalEventos
    wsDash.Range("E3").Value = "Riesgos Medios:"
    wsDash.Range("F3").Value = conteoMedio
    wsDash.Range("E4").Value = "Riesgos Altos:"
    wsDash.Range("F4").Value = conteoAlto
    wsDash.Range("E5").Value = "Riesgos Críticos:"
    wsDash.Range("F5").Value = conteoCritico
    wsDash.Range("E6").Value = "Última actualización:"
    wsDash.Range("F6").Value = Now

    ' Termómetro de riesgo
    If conteoCritico > 0 Then
        wsDash.Range("E8").Value = "RIESGO GLOBAL: CRÍTICO"
        wsDash.Range("E8").Interior.Color = RGB(255, 0, 0)
    ElseIf conteoAlto > 0 Then
        wsDash.Range("E8").Value = "RIESGO GLOBAL: ALTO"
        wsDash.Range("E8").Interior.Color = RGB(255, 100, 0)
    ElseIf conteoMedio > 0 Then
        wsDash.Range("E8").Value = "RIESGO GLOBAL: MEDIO"
        wsDash.Range("E8").Interior.Color = RGB(255, 255, 0)
    Else
        wsDash.Range("E8").Value = "RIESGO GLOBAL: BAJO"
        wsDash.Range("E8").Interior.Color = RGB(0, 255, 0)
    End If
    wsDash.Range("E8").Font.Bold = True
    wsDash.Range("E8").Font.Size = 14
    wsDash.Range("E8").HorizontalAlignment = xlCenter
End Sub

