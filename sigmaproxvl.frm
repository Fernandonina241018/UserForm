VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sigmaproxvl 
   Caption         =   "SigmaProXVL2.0"
   ClientHeight    =   12420
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   18765
   OleObjectBlob   =   "sigmaproxvl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sigmaproxvl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'#Else
'    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
'        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
'        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Const GWL_STYLE = -16
Const WS_SYSMENU = &H80000

' =============================================================================
' USERFORM MODERNO - ANÁLISIS ESTADÍSTICO CON AUTO-POSICIONAMIENTO
' =============================================================================
' Diseño inspirado en ClickUp con posicionamiento automático desde código
' =============================================================================

Option Explicit

' Variables para arrastrar el formulario
Private isDragging As Boolean
Private dragX As Long
Private dragY As Long

' Constantes de diseño
Private Const MARGEN_LATERAL As Long = 20
Private Const MARGEN_TOP As Long = 70
Private Const ESPACIADO_VERTICAL As Long = 12
Private Const ALTURA_CONTROL As Long = 26
Private Const ALTURA_FRAME As Long = 80
Private Const ANCHO_LABEL As Long = 140
Private Const ANCHO_COMBOBOX As Long = 200
Public RangoDetectado As Range

Private Sub btnRandodeColumnas_Click()
    Call ObtenerRango(txtRangodeColumnas)
End Sub

Private Sub btnRandodeDatosHM_Click()
    If txtRango <> "" Then
        txtRangodeDatosHM = txtRango.Value
        Call ObtenerRango(txtRangodeDatosHM)
    End If
End Sub

Private Sub btnRandodeFilas_Click()
    Call ObtenerRango(txtRangodeFilas)
End Sub

Private Sub btnRandoDestino_Click()
    Call ObtenerRango(txtRangoDestino)
End Sub

Private Sub btnVariableX_Click()
    Call ObtenerRango(txtVariableX)
End Sub

Private Sub btnVariableY_Click()
    Call ObtenerRango(txtVariableY)
End Sub

Private Sub cboHojaTrabajo_Change()
    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    Dim ws As Worksheet
    Dim nombreHoja As String
    
    nombreHoja = Trim(cboHojaTrabajo.Value)
    If nombreHoja = "" Then GoTo SafeExit
    
    ' Validar existencia de hoja
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "La hoja '" & nombreHoja & "' no existe.", vbExclamation
    ElseIf ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
        Debug.Print "La hoja '" & nombreHoja & "' está oculta.", vbExclamation
    Else
        ws.Activate
    End If
    
SafeExit:
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description, vbExclamation
    Resume SafeExit
End Sub

Private Sub UserForm_Activate()
    Me.Width = 950
    Me.Height = 650

    Me.Left = (Application.Width - Me.Width) / 2   ' centrar horizontal
    Me.Top = (Application.Height - Me.Height) / 2   ' centrar vertical
End Sub


Private Sub cboLibrodeTrabajo_Change()
    Call ActualizarHojas
End Sub

Private Sub cboLimiteInferior_Change()

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub cboRangodeDatosHM_Change()

End Sub

Private Sub cboModoAnalisis_Change()

End Sub

Private Sub chkPage1_Click()
    If chkPage1.Value = True Then
        chkPage2.Value = False
        chkPage3.Value = False
        chkPage4.Value = False
        MultiPageDatos.Value = 0
    End If
End Sub

Private Sub chkPage2_Click()
    If Me.chkPage2.Value = True Then
        Me.chkPage1.Value = False
        Me.chkPage3.Value = False
        Me.chkPage4.Value = False
        Me.MultiPageDatos.Value = 1
    End If
End Sub

Private Sub chkPage3_Click()
    If Me.chkPage3.Value = True Then
        Me.chkPage2.Value = False
        Me.chkPage1.Value = False
        Me.chkPage4.Value = False
        Me.MultiPageDatos.Value = 2
    End If
End Sub

Private Sub chkPage4_Click()
    If Me.chkPage4.Value = True Then
        Me.chkPage2.Value = False
        Me.chkPage1.Value = False
        Me.chkPage3.Value = False
        Me.MultiPageDatos.Value = 3
    End If
End Sub

Private Sub frameOpcionesAnalisis_Click()

End Sub

Private Sub frameCalculosF0_Click()

End Sub

Private Sub frameOpcionesAnalisis0_Click()

End Sub

Private Sub frameOpcionesAnalisis1_Click()

End Sub

Private Sub lblTiempoInicio_Click()

End Sub

Private Sub btnBorrarSheets_Click()
    Dim respuesta As VbMsgBoxResult

    respuesta = MsgBox("¿Estás seguro de que deseas borrar todas las hojas excepto la primera?", vbYesNo + vbQuestion, "Confirmar borrado")

    If respuesta = vbYes Then
        Call BorrarHojasMenosLaPrimera_Activo
        MsgBox "Las hojas han sido borradas correctamente.", vbInformation, "Proceso completado"
    Else
        MsgBox "Operación cancelada.", vbExclamation, "Cancelado"
    End If
End Sub

Private Sub TextBoxRango_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validar cuando el usuario sale del TextBox
    
    If txtRango.Value <> "" Then
        If Not ValidarRango(txtRango.Value) Then
            MsgBox "La dirección de rango ingresada no es válida.", vbExclamation
            Cancel = True ' Evitar que salga del TextBox
        End If
    End If
    
End Sub

Private Sub lstDatosColumnas_Click()
    Dim columnaSeleccionada As Long
    Dim letraColumna As String
    Dim filaInicio As Long, filaFin As Long
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim rangoAutomatico As Range

    Set wb = Workbooks(cboLibrodeTrabajo.Value)
    Set ws = wb.Worksheets(cboHojaTrabajo.Value)
    Set rangoAutomatico = ws.UsedRange

    filaInicio = rangoAutomatico.Row
    filaFin = filaInicio + rangoAutomatico.Rows.count - 1
    columnaSeleccionada = rangoAutomatico.Columns(lstDatosColumnas.ListIndex + 1).Column

    letraColumna = Split(Cells(1, columnaSeleccionada).Address(True, False), "$")(0)

    TextBoxRango.Value = "'" & ws.Name & "'!" & "$" & letraColumna & "$" & filaInicio & ":$" & letraColumna & "$" & filaFin
End Sub

Private Sub txtRango_Change()
    ' Opcional: Formatear el texto mientras se escribe
    
    ' Convertir a mayúsculas automáticamente
    Dim PosicionCursor As Long
    PosicionCursor = txtRango.SelStart
    
    txtRango.Value = UCase(txtRango.Value)
    txtRango.SelStart = PosicionCursor
End Sub

' =============================================================================
' EVENTO: UserForm_Initialize
' =============================================================================
Private Sub UserForm_Initialize()
    Dim wb As Workbook
    
    ' Limpiar controles
    lstDatosColumnas.Clear
    cboLibrodeTrabajo.Clear
    
    ' Cargar libros de trabajo
    For Each wb In Application.Workbooks
        ' NO agregar PERSONAL.XLSB al combo
        If wb.Name <> "PERSONAL.XLSB" And wb.Name <> "PERSONAL.XLS" Then
            cboLibrodeTrabajo.AddItem wb.Name
        End If
    Next wb
    
    ' Seleccionar el primer libro (que NO será PERSONAL)
    If cboLibrodeTrabajo.ListCount > 0 Then
        cboLibrodeTrabajo.ListIndex = 0
        ' Solo llamar ActualizarHojas UNA VEZ aquí
        Call ActualizarHojas
    End If
    
    Dim lStyle As Long
    lStyle = lStyle And Not WS_SYSMENU
    
    Me.BackColor = RGB(255, 255, 255)
    
    Call ConfigurarTemaModerno
    Call InicializarComboBoxes
    Call InicializarCheckBoxes
    Call EstilizarBotones
    
    MultiPageDatos.Value = 0
    chkPage1.Value = True
    lblNombredeUsuario.Caption = "USER NAME: " & Environ("Username")
    lblNombrePC.Caption = "PC NAME: " & Environ("ComputerName")
    Me.btnSeleccionarRango.SetFocus
End Sub



' =============================================================================
' CONFIGURAR TEMA MODERNO
' =============================================================================
Private Sub ConfigurarTemaModerno()
    
    ' ===== BARRA DE TÍTULO =====
    With lblTitulo
        .Caption = "Análisis Estadístico"
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
    End With


    With lblSubtitulo
        .Caption = "Herramienta profesional de análisis de datos y control de calidad"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(108, 117, 125)
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' ===== FRAME 1: SELECCIÓN DE DATOS =====
    With frameSeleccionDatos
        .Caption = "Selección de Datos"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With FrameSeccionTitulo
        .Caption = " "
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With frameCargadeDatos
        .Caption = " "
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With FrameSeccionMedia
        .Caption = "Opciones Extras"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With lstDatosColumnas
        '.Caption = "Opciones Extras"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With lblmodo
        .Caption = "Modo de Análisis:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblDescripcionEstadistica
        .Caption = "Carga de Datos:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblLibro
        .Caption = "Libro de Trabajo:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblHoja
        .Caption = "Hoja de Trabajo:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblRangoDatos
        .Caption = "Rango de datos:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblRangodeColumnas
        .Caption = "Rango de columnas:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With LblRangodeDatosHM
        .Caption = "Rango de datos HeatMap:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblVariableX
        .Caption = "VariableX:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblVariableY
        .Caption = "VariableY:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblRangodeFilas
        .Caption = "Rango de datos Fila:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblRangoDestino
        .Caption = "Rango Destina:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblNombredeUsuario
        '.Caption = ""
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
    End With
    
    With lblNombrePC
        '.Caption = ""
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
    End With
    
    With txtRango
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With txtRangodeColumnas
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With txtRangodeDatosHM
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With txtRangodeFilas
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With txtRangoDestino
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With txtVariableX
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With txtVariableY
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    ' ===== FRAME 2: LÍMITES Y PARÁMETROS =====
    With frameLimitesParametros
        .Caption = "Límites y Parámetros"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    Call EstilizarLabel(lblLimiteSuperior, "Límite Superior:")
    Call EstilizarLabel(lblLimiteInferior, "Límite Inferior:")
    Call EstilizarLabel(lblExpectativa, "Expectativa Matemática:")
    
    Call EstilizarComboBox(cboLimiteSuperior)
    Call EstilizarComboBox(cboLimiteInferior)
    Call EstilizarComboBox(cboExpectativa)
    Call EstilizarComboBox(cboModoAnalisis)
    Call EstilizarComboBox(cboHojaTrabajo)
    Call EstilizarComboBox(cboLibrodeTrabajo)
    
    ' ===== FRAME 3: CÁLCULOS F0 =====
    With frameCalculosF0
        .Caption = "Apartado de Opciones Principales"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    Call EstilizarLabel(lblTiempoInicio, "Tiempo Inicio:")
    Call EstilizarLabel(lblTiempoFinal, "Tiempo Final:")
    
    Call EstilizarComboBox(cboTiempoInicio)
    Call EstilizarComboBox(cboTiempoFinal)
    
    ' ===== FRAME 4: OPCIONES DE ANÁLISIS =====
    With frameOpcionesAnalisis0
        If Me.txtRango <> "" Then
            .Caption = "Opciones de Análisis Estandar [Cargado]"
        Else
            .Caption = "Opciones de Análisis Estandar [Libre]"
        End If
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With frameOpcionesAnalisis1
        .Caption = "Opciones de Análisis HeatMap"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With frameOpcionesAnalisis2
        .Caption = "Opciones de Análisis Regresión Lineal"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With frameOpcionesAnalisis3
        .Caption = "Opciones de Análisis (4)"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    
    With FrameBtn
        .Caption = "Opciones de Ejecución"
        .BackColor = RGB(248, 249, 250)
        .BorderColor = RGB(222, 226, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(33, 37, 41)
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    With MultiPageDatos
        .BackColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Style = fmTabStyleNone
    End With


    ' Personalizar contenido dentro de cada página
    Dim i As Integer
    For i = 0 To MultiPageDatos.Pages.count - 1
        ' Si hay un Frame en la página, cambiar su color
        Dim ctrl As Control
        For Each ctrl In MultiPageDatos.Pages(i).Controls
            If TypeName(ctrl) = "Frame" Then
                ctrl.BackColor = RGB(248, 249, 250)
                ctrl.BorderColor = RGB(222, 226, 230)
            End If
        Next ctrl
    Next i

    Call EstilizarCheckBox(chkCorrelacion, "Análisis de Correlación")
    Call EstilizarCheckBox(chkCapacidadProceso, "Capacidad de Proceso (Cpk)")
    Call EstilizarCheckBox(chkGraficos, "Generar Gráficos de Control")
    Call EstilizarCheckBox(chkOutliers, "Detectar Outliers (IQR)")
    Call EstilizarCheckBox(chkPage1, "Vizualizar Análisis Estandar")
    Call EstilizarCheckBox(chkPage2, "Vizualizar Análisis HeatMap")
    Call EstilizarCheckBox(chkPage3, "Vizualizar Análisis Regresión Lineal")
    Call EstilizarCheckBox(chkPage4, "Vizualizar Página (4)")
    Call EstilizarCheckBox(chkHeatMap, "Generar Gráfico HeatMap")
    Call EstilizarCheckBox(chkRegresionLineal, "Regresión Lineal")
    Call EstilizarCheckBox(Me.chkAutoRango, "Cargar Datos")
    
End Sub

' =============================================================================
' INICIALIZAR COMBOBOXES
' =============================================================================
Private Sub InicializarComboBoxes()
    
    With cboLimiteSuperior
        .Clear
        .AddItem "25.0"
        .AddItem "30.0"
        .AddItem "37.0"
        .AddItem "121.0"
        .AddItem "150.0"
        .AddItem "123.4"
        .AddItem "150.0"
        .ListIndex = 0
    End With
    
    With cboLimiteInferior
        .Clear
        .AddItem "2.0"
        .AddItem "8.0"
        .AddItem "15.0"
        .AddItem "20.0"
        .AddItem "25.0"
        .AddItem "118.5"
        .ListIndex = 0
    End With
    
    With cboExpectativa
        .Clear
        .AddItem "0"
        .AddItem "100"
        .AddItem "121"
        .AddItem "30"
        .AddItem "80"
        .AddItem "75"
        .AddItem "65"
        .AddItem "25"
        .AddItem "5"
        .Text = "0"
    End With
    
    With cboTiempoInicio
        .Clear
        .AddItem "00:00"
        .AddItem "08:00"
        .AddItem "12:00"
        .Text = "00:00"
    End With
    
    With cboTiempoFinal
        .Clear
        .AddItem "01:00"
        .AddItem "02:00"
        .AddItem "04:00"
        .AddItem "08:00"
        .Text = "01:00"
    End With
    
    With cboModoAnalisis
        .Clear
        .AddItem "Análisis"
        .AddItem "Esterilización"
        .AddItem "Mapeo"
        .AddItem "Calificación"
        .Text = "Análisis"
    End With
    
End Sub

' =============================================================================
' INICIALIZAR CHECKBOXES
' =============================================================================
Private Sub InicializarCheckBoxes()
    chkCorrelacion.Value = False
    chkCapacidadProceso.Value = False
    chkGraficos.Value = True
    chkOutliers.Value = True
    chkPage1.Value = False
    chkPage2.Value = False
    chkPage3.Value = False
    chkPage4.Value = False
    Me.chkRegresionLineal.Value = False
    Me.chkHeatMap.Value = False
End Sub

' =============================================================================
' ESTILIZAR LABEL
' =============================================================================
Private Sub EstilizarLabel(lbl As MSForms.Label, texto As String)
    With lbl
        .Caption = texto
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(73, 80, 87)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
    End With
End Sub

' =============================================================================
' ESTILIZAR COMBOBOX
' =============================================================================
Private Sub EstilizarComboBox(cbo As MSForms.ComboBox)
    With cbo
        ' Colores base Material Design
        .BackColor = RGB(255, 255, 255)          ' Blanco puro
        .ForeColor = RGB(33, 33, 33)             ' Texto casi negro
        .BorderColor = RGB(189, 189, 189)        ' Borde gris claro
        .BorderStyle = fmBorderStyleSingle
        
        ' Tipografía moderna
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = False
        
        ' Alineación y espaciado
        .TextAlign = fmTextAlignCenter
        .SpecialEffect = fmSpecialEffectFlat
        
        ' Dimensiones recomendadas
        .Height = 25
        
        ' Efecto visual de profundidad (simulado con color de fondo)
        ' Al hacer hover o focus, cambiará en los eventos
    End With
End Sub

' =============================================================================
' ESTILIZAR CHECKBOX
' =============================================================================
Private Sub EstilizarCheckBox(chk As MSForms.CheckBox, texto As String)
    With chk
        .Caption = texto
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .ForeColor = RGB(33, 37, 41)
        .BackStyle = fmBackStyleTransparent
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' =============================================================================
' ESTILIZAR BOTONES
' =============================================================================
Private Sub EstilizarBotones()
    
    With btnAnalizar
        .Caption = "Ejecutar Análisis"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With CommandButton1
        .Caption = "Disp 1"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With CommandButton2
        .Caption = "Disp 2"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    
    With btnLimpiar
        .Caption = "Limpiar Campos"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnRandodeDatosHM
        .Caption = "Datos"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnRandodeColumnas
        .Caption = "Col. Date"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnRandodeFilas
        .Caption = "Row. Date"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnVariableX
        .Caption = "Var. X"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnVariableY
        .Caption = "Var. Y"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnRandoDestino
        .Caption = "Destino"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnAyuda
        .Caption = "Help"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnSeleccionarRango
        .Caption = "Datos"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
    With btnBorrarSheets
        .Caption = "Delete"
        .BackColor = RGB(189, 201, 200)
        .ForeColor = RGB(0, 0, 0)
        .Font.Size = 9
        .Font.Bold = True
        .TakeFocusOnClick = False
    End With
    
End Sub

' =============================================================================
' EVENTOS DE BOTONES
' =============================================================================

Private Sub btnAnalizar_Click()
    
    If Trim(txtRango.Text) = "" Then
        MsgBox "Por favor, selecciona un rango de datos para analizar.", _
               vbExclamation, "Campo Requerido"
        txtRango.SetFocus
        Exit Sub
    End If
    
    If Not ValidarRango(txtRango.Text) Then
        MsgBox "El formato del rango no es válido." & vbCrLf & vbCrLf & _
               "Formato correcto: A1:D100", _
               vbExclamation, "Rango Inválido"
        Exit Sub
    End If
    
    Me.Hide
    If chkGraficos.Value = True Or chkOutliers.Value = True Then
        MostrarAnalisisVerticalEnHoja
    Else
        MsgBox "Deben estar seleccionados "
    End If
    
    Unload Me
    
    Dim RangoDatos As Range
    Dim rangoDestino As Range
    Dim rangoFilas As Range
    Dim rangoColumnas As Range
    Dim direccion As String
    
    If chkRegresionLineal.Value = True Then
        ' Validar rango de datos
        If txtRangodeDatosHM.Value = "" Then
            MsgBox "Debe seleccionar el rango de datos.", vbExclamation, "Error"
            Exit Sub
        End If
    
        ' Validar rango destino
        If txtRangoDestino.Value = "" Then
            MsgBox "Debe seleccionar la celda destino.", vbExclamation, "Error"
            Exit Sub
        End If
    
        On Error Resume Next
        Set RangoDatos = Range(txtRangodeDatosHM.Value)
        Set rangoDestino = Range(txtRangoDestino.Value)
    
        If txtRangodeFilas.Value <> "" Then
            Set rangoFilas = Range(txtRangodeFilas.Value)
        End If
    
        If txtRangodeColumnas.Value <> "" Then
            Set rangoColumnas = Range(txtRangodeColumnas.Value)
        End If
        On Error GoTo 0
    
        If RangoDatos Is Nothing Then
            MsgBox "El rango de datos no es válido.", vbExclamation, "Error"
            Exit Sub
        End If
    
        If rangoDestino Is Nothing Then
            MsgBox "El rango destino no es válido.", vbExclamation, "Error"
            Exit Sub
        End If
    End If
    
    ' Tómate un momento para asegurar que no hay errores de sintaxis anidados de la corrección anterior.
    If chkHeatMap.Value = True Then
        
        ' Usamos Range() sobre el valor de los RefEdit para obtener el objeto Range.
        Call GenerarMapaCalorDesdeRangos(Range(txtRangoDestino.Value), _
            Range(txtRangoDestino.Value), _
            Range(txtRangodeFilas.Value), _
            Range(txtRangodeColumnas.Value), _
            chkHeatMap.Value)  ' Este tipo de dato ya es correcto (Boolean)
    End If
    
    ' Este código se ejecuta DESPUÉS de cerrar el formulario
    If chkRegresionLineal.Value = True Then
        Call IniciarRegresionConRefEdit
    End If
    
    On Error GoTo manejaErr

    direccion = txtRangodeDatosHM.Value   ' Ej.: [Book1]Sheet1!$C$1:$F$5

    Debug.Print "Resultados escritos en la columna 7 (G).", vbInformation
    Exit Sub

manejaErr:
    MsgBox "Error: " & Err.Description, vbExclamation
    
End Sub

Private Sub btnSeleccionarRango_Click()
    Call ObtenerRango(txtRango)
End Sub

Private Sub btnLimpiar_Click()
    
    txtRango.Text = ""
    cboLimiteSuperior.ListIndex = 0
    cboLimiteInferior.ListIndex = 0
    cboExpectativa.Text = "0"
    cboTiempoInicio.Text = "00:00"
    cboTiempoFinal.Text = "01:00"
    
    chkCorrelacion.Value = False
    chkCapacidadProceso.Value = False
    chkGraficos.Value = True
    chkOutliers.Value = True
    
    txtRango.SetFocus
    
End Sub

'Private Sub btnCerrar_Click()
'    Unload Me
'End Sub

Private Sub btnAyuda_Click()
    
    Dim ayuda As String
    ayuda = "?? AYUDA - ANÁLISIS ESTADÍSTICO" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    ayuda = ayuda & "1SELECCIÓN DE DATOS:" & vbCrLf
    ayuda = ayuda & "   • Haz clic en ?? para seleccionar el rango de datos" & vbCrLf
    ayuda = ayuda & "   • Formato: A1:D100 (incluye encabezados)" & vbCrLf & vbCrLf
    
    ayuda = ayuda & "2LÍMITES Y PARÁMETROS:" & vbCrLf
    ayuda = ayuda & "   • Define límites superior e inferior para control" & vbCrLf
    ayuda = ayuda & "   • Expectativa: valor objetivo esperado" & vbCrLf & vbCrLf
    
    ayuda = ayuda & "3OPCIONES DE ANÁLISIS:" & vbCrLf
    ayuda = ayuda & "   ? Correlación: Analiza relación entre variables" & vbCrLf
    ayuda = ayuda & "   ? Cpk: Capacidad del proceso" & vbCrLf
    ayuda = ayuda & "   ? Gráficos: Genera visualizaciones" & vbCrLf
    ayuda = ayuda & "   ? Outliers: Detecta valores atípicos" & vbCrLf & vbCrLf
    
    ayuda = ayuda & "4EJECUTAR:" & vbCrLf
    ayuda = ayuda & "   • Clic en '? Ejecutar Análisis' para iniciar" & vbCrLf & vbCrLf
    
    ayuda = ayuda & "?? Los resultados se generarán en una nueva hoja."
    
    MsgBox ayuda, vbInformation, "Ayuda - Análisis Estadístico"
    
End Sub

' =============================================================================
' VALIDACIÓN
' =============================================================================
Private Function ValidarRango(rangoTexto As String) As Boolean
    
    Dim rng As Range
    
    On Error Resume Next
    Set rng = Range(rangoTexto)
    On Error GoTo 0
    
    ValidarRango = Not rng Is Nothing
    
End Function

' =============================================================================
' FUNCIONALIDAD DE ARRASTRE
' =============================================================================
Private Sub lblTitulo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal Y As Single)
    isDragging = True
    dragX = x
    dragY = Y
End Sub

Private Sub lblTitulo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal Y As Single)
    If isDragging Then
        Me.Left = Me.Left + (x - dragX)
        Me.Top = Me.Top + (Y - dragY)
    End If
End Sub

Private Sub lblTitulo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                              ByVal x As Single, ByVal Y As Single)
    isDragging = False
End Sub

' =============================================================================
' PROPIEDADES PÚBLICAS PARA COMPATIBILIDAD
' =============================================================================

' Para mantener compatibilidad con código existente que usa:
' sigmaproxvl.TextBoxRango.Text
Public Property Get TextBoxRango() As MSForms.TextBox
    Set TextBoxRango = txtRango
End Property

Public Property Get ComboBoxLimiteSuperior() As String
    ComboBoxLimiteSuperior = cboLimiteSuperior.Text
End Property

Public Property Get ComboBoxLimiteInferior() As String
    ComboBoxLimiteInferior = cboLimiteInferior.Text
End Property

Public Property Get ComboBoxExperanzaMath() As MSForms.ComboBox
    Set ComboBoxExperanzaMath = cboExpectativa
End Property

Public Property Get ComboBoxTiempoFinal() As MSForms.ComboBox
    Set ComboBoxTiempoFinal = cboTiempoFinal
End Property

Public Property Get ComboBoxTiempodeInicio() As MSForms.ComboBox
    Set ComboBoxTiempodeInicio = cboTiempoInicio
End Property

Public Property Get CheckBoxCorrelacion() As MSForms.CheckBox
    Set CheckBoxCorrelacion = chkCorrelacion
End Property

Public Property Get CheckBoxCapacidadProceso() As MSForms.CheckBox
    Set CheckBoxCapacidadProceso = chkCapacidadProceso
End Property

Private Sub ActualizarHojas()
    Dim wb As Workbook
    Dim i As Long
    cboHojaTrabajo.Clear
    On Error Resume Next
    Set wb = Application.Workbooks(cboLibrodeTrabajo.Value)
    On Error GoTo 0
    If Not wb Is Nothing Then
        For i = 1 To wb.Sheets.count
            cboHojaTrabajo.AddItem wb.Sheets(i).Name
        Next i
        If cboHojaTrabajo.ListCount > 0 Then
            cboHojaTrabajo.ListIndex = 0
        End If
    End If
End Sub

Sub AlinearControlesVerticalmente()
    Dim controles As Variant
    Dim i As Integer
    Dim espacio As Integer
    Dim posicionInicialTop As Integer

    ' Lista de nombres de controles
    controles = Array("lblLimiteSuperior", "cboLimiteSuperior", _
                        "lblLimiteInferior", "cboLimiteInferior", _
                        "lblTiempoInicio", "cboTiempoInicio", _
                        "lblTiempoFinal", "cboTiempoFinal", _
                        "lblExpectativa", "cboExpectativa")

    ' Espacio entre controles (en píxeles)
    espacio = 10

    ' Posición inicial desde arriba
    posicionInicialTop = 20

    ' Recorre los controles y los posiciona
    For i = 0 To UBound(controles)
        With Me.Controls(controles(i))
            .Top = posicionInicialTop
            posicionInicialTop = .Top + .Height + espacio
        End With
    Next i
End Sub

' Procedimiento para mostrar resultados (usar desde UserForm)
Sub MostrarAnalisis()
    Dim rango As Range
    Dim stats As Estadisticas
    
    ' Obtener rango desde RefEdit (cambiar "RefEdit1" por el nombre de tu control)
    On Error Resume Next
    Set rango = Range(txtRango.Text)
    On Error GoTo 0
    
    If rango Is Nothing Then
        Debug.Print "Rango inválido"
        Exit Sub
    End If
    
    ' Calcular estadísticas
    stats = AnalizarCoordenadas(rango)
    
    ' Mostrar resultados (ejemplo en debug.print)
    With stats
        Debug.Print _
            "Count: " & .count & vbCrLf & _
            "Promedio: " & Format(.promedio, "0.0000") & vbCrLf & _
            "Desviación Estándar: " & Format(.desviacionEstandar, "0.0000") & vbCrLf & _
            "RSD: " & Format(.RSD, "0.0000") & "%" & vbCrLf & _
            "Máximo: " & Format(.maximo, "0.0000") & vbCrLf & _
            "Mínimo: " & Format(.minimo, "0.0000")
    End With
End Sub

'-------------------------------------------------------------------------------
' EVENTO: chkAutoRango_Click
' PROPÓSITO: Detectar automáticamente el rango de datos y cargar encabezados
' VERSIÓN: 2.0 (Optimizada)
'-------------------------------------------------------------------------------

Private Sub chkAutoRango_Click()
    
    If chkAutoRango.Value = True Then
        '-----------------------------------------------------------------------
        ' MODO AUTOMÁTICO: Detectar rango y cargar encabezados
        '-----------------------------------------------------------------------
        Call ActivarModoAutomatico
    Else
        '-----------------------------------------------------------------------
        ' MODO MANUAL: Restaurar controles
        '-----------------------------------------------------------------------
        Call ActivarModoManual
    End If
    
End Sub

'-------------------------------------------------------------------------------
' PROCEDIMIENTO: ActivarModoAutomatico
' PROPÓSITO: Lógica para modo automático (refactorizada)
'-------------------------------------------------------------------------------

Private Sub ActivarModoAutomatico()
    
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rangoAutomatico As Range
    Dim encabezado As Range
    Dim celda As Range
    Dim valorCelda As String
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    
    '---------------------------------------------------------------------------
    ' VALIDACIÓN 1: Libro de trabajo seleccionado
    '---------------------------------------------------------------------------
    If cboLibrodeTrabajo.ListIndex = -1 Then
        MsgBox "? Por favor, selecciona un libro de trabajo primero.", _
               vbExclamation, "Libro no seleccionado"
        Me.chkAutoRango.Value = False
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------
    ' OBTENER LIBRO DE TRABAJO
    '---------------------------------------------------------------------------
    On Error Resume Next
    Set wb = Workbooks(cboLibrodeTrabajo.Value)
    If Err.Number <> 0 Then
        MsgBox "? Error al acceder al libro: " & Err.Description, _
               vbCritical, "Error de Acceso"
        Me.chkAutoRango.Value = False
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    If wb Is Nothing Then
        MsgBox "? No se pudo acceder al libro seleccionado.", _
               vbExclamation, "Libro Inaccesible"
        Me.chkAutoRango.Value = False
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------
    ' VALIDACIÓN 2: Hoja de trabajo seleccionada
    '---------------------------------------------------------------------------
    If Me.cboHojaTrabajo.ListIndex = -1 Then
        MsgBox "? Por favor, selecciona una hoja de trabajo primero.", _
               vbExclamation, "Hoja no seleccionada"
        Me.chkAutoRango.Value = False
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------
    ' OBTENER HOJA DE TRABAJO
    '---------------------------------------------------------------------------
    On Error Resume Next
    Set ws = wb.Worksheets(cboHojaTrabajo.Value)
    If Err.Number <> 0 Then
        MsgBox "? Error al acceder a la hoja: " & Err.Description, _
               vbCritical, "Error de Acceso"
        Me.chkAutoRango.Value = False
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "? No se pudo acceder a la hoja seleccionada.", _
               vbExclamation, "Hoja Inaccesible"
        Me.chkAutoRango.Value = False
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------
    ' DETECTAR RANGO CON DATOS REALES (Método mejorado)
    '---------------------------------------------------------------------------
    ' Encontrar última fila con datos
    ultimaFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' Encontrar última columna con datos
    ultimaColumna = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    ' Validar que hay datos
    If ultimaFila < 1 Or ultimaColumna < 1 Then
        MsgBox "? No se detectaron datos en la hoja seleccionada.", _
               vbExclamation, "Sin Datos"
        Me.chkAutoRango.Value = False
        Exit Sub
    End If
    
    ' Crear rango automático optimizado
    Set rangoAutomatico = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila, ultimaColumna))
    
    '---------------------------------------------------------------------------
    ' MOSTRAR RANGO EN TEXTBOX
    '---------------------------------------------------------------------------
    txtRango.Text = "'" & ws.Name & "'!" & rangoAutomatico.Address
     'TextBoxRango.Value = "'" & ws.Name & "'!" & "$" & letraColumna & "$" & filaInicio & ":$" & letraColumna & "$" & filaFin
    
'    '---------------------------------------------------------------------------
'    ' CARGAR ENCABEZADOS EN LISTBOX (Método mejorado)
'    '---------------------------------------------------------------------------
'    lstDatosColumnas.Clear
'
'    Set encabezado = rangoAutomatico.Rows(1)
'
'    For Each celda In encabezado.Cells
'        ' Limpiar y validar valor de encabezado
'        valorCelda = Trim(CStr(celda.Value))
'
'        If Len(valorCelda) > 0 Then
'            ' Agregar encabezado con referencia de columna
'            lstDatosColumnas.AddItem valorCelda & " [Col " & _
'                                     Split(celda.Address, "$")(1) & "]"
'        End If
'    Next celda
'
'    ' Validar que se encontraron encabezados
'    If lstDatosColumnas.ListCount = 0 Then
'        MsgBox "? No se encontraron encabezados válidos en la primera fila.", _
'               vbExclamation, "Sin Encabezados"
'        Me.chkAutoRango.Value = False
'        Exit Sub
'    End If
    '
    ' Suponiendo que las variables rangoAutomatico, celda, valorCelda, lstDatosColumnas
    ' y el rango de datos existen y están correctamente declaradas/establecidas.
    
    '---------------------------------------------------------------------------
    ' CARGAR ENCABEZADOS EN LISTBOX (Método mejorado)
    '---------------------------------------------------------------------------
    lstDatosColumnas.Clear
    lstDatosColumnas.ColumnCount = rangoAutomatico.Columns.count ' <-- 1. CRUCIAL: Definir el número de columnas del ListBox
    Dim ColIndex As Long ' Para capturar el índice de la columna
    Dim FilaActual As Range ' Para iterar sobre cada fila de datos
    
    Set encabezado = rangoAutomatico.Rows(1)
    
    ' === PASO 1: Cargar Encabezados ===
    For Each celda In encabezado.Cells
        ColIndex = celda.Column - rangoAutomatico.Column ' Calcula el índice basado en 0
        
        ' Limpiar y validar valor de encabezado
        valorCelda = Trim(CStr(celda.Value))
        
        If Len(valorCelda) > 0 Then
            ' Agregar encabezado con referencia de columna (Solo se agrega la primera vez)
            ' Usamos AddItem en la primera columna, luego List(index, ColIndex)
            lstDatosColumnas.AddItem valorCelda ' Agrega el primer elemento a la fila
        End If
    Next celda
    
    ' Validar que se encontraron encabezados
    If lstDatosColumnas.ListCount = 0 Then
        MsgBox "? No se encontraron encabezados válidos en la primera fila.", vbExclamation, "Sin Encabezados"
        Me.chkAutoRango.Value = False
        Exit Sub
    End If
    
    ' === PASO 2: Cargar Datos (Excluyendo la primera fila) ===
    
    ' Recorrer todas las filas del rango, empezando desde la segunda fila (Rows(2))
    For Each FilaActual In rangoAutomatico.Rows
        If FilaActual.Row > encabezado.Row Then ' Asegura que estamos en una fila de datos
            
            Dim i As Long
            Dim EsPrimeraColumna As Boolean
            EsPrimeraColumna = True
            
            ' Recorrer cada celda dentro de la fila actual
            For i = 1 To FilaActual.Cells.count
                valorCelda = FilaActual.Cells(i).Value ' Obtener el valor de la celda
                
                If EsPrimeraColumna Then
                    ' Primera columna: Usamos AddItem (Crea la nueva fila en el ListBox)
                    lstDatosColumnas.AddItem valorCelda
                    EsPrimeraColumna = False
                Else
                    ' Columnas subsiguientes: Usamos la propiedad .List para agregar el ítem
                    ' a la columna 'i-1' de la fila recién agregada (.ListCount - 1)
                    lstDatosColumnas.List(lstDatosColumnas.ListCount - 1, i - 1) = valorCelda
                End If
            Next i
        End If
    Next FilaActual
    '---------------------------------------------------------------------------
    ' CONFIGURAR CONTROLES PARA MODO AUTOMÁTICO
    '---------------------------------------------------------------------------
    txtRango.Enabled = True
    btnSeleccionarRango.Enabled = False
    
    ' Feedback visual positivo
    Me.chkAutoRango.ForeColor = RGB(0, 128, 0)  ' Verde
    
    Exit Sub

ErrorHandler:
    '---------------------------------------------------------------------------
    ' MANEJO CENTRALIZADO DE ERRORES
    '---------------------------------------------------------------------------
    MsgBox "? Error inesperado: " & Err.Description & vbCrLf & _
           "Número: " & Err.Number, vbCritical, "Error en Modo Automático"
    Me.chkAutoRango.Value = False
    Me.chkAutoRango.ForeColor = RGB(0, 0, 0)  ' Negro
    Err.Clear
    
End Sub

'-------------------------------------------------------------------------------
' PROCEDIMIENTO: ActivarModoManual
' PROPÓSITO: Restaurar controles para modo manual
'-------------------------------------------------------------------------------

Private Sub ActivarModoManual()
    
    ' Limpiar controles
    txtRango.Text = ""
    lstDatosColumnas.Clear
    
    ' Habilitar controles manuales
    txtRango.Enabled = True
    btnSeleccionarRango.Enabled = True
    
    ' Restaurar color del checkbox
    Me.chkAutoRango.ForeColor = RGB(0, 0, 0)  ' Negro
    
    ' Establecer foco
    txtRango.SetFocus
    
End Sub
