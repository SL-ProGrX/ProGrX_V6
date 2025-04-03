VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Vacaciones_Acumuladas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Cargado Vacaciones Acumuladas"
   ClientHeight    =   9528
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   14268
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9528
   ScaleWidth      =   14268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   312
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   5532
      _Version        =   1310720
      _ExtentX        =   9758
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   672
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   5532
      _Version        =   1310720
      _ExtentX        =   9758
      _ExtentY        =   1185
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   1560
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmRH_Vacaciones_Acumuladas.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   1
      Left            =   7440
      TabIndex        =   4
      Top             =   1560
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmRH_Vacaciones_Acumuladas.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   2
      Left            =   7920
      TabIndex        =   5
      Top             =   1560
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmRH_Vacaciones_Acumuladas.frx":0E19
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6852
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   14052
      _Version        =   524288
      _ExtentX        =   24786
      _ExtentY        =   12086
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmRH_Vacaciones_Acumuladas.frx":1532
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpActualiza 
      Height          =   312
      Left            =   11640
      TabIndex        =   10
      Top             =   1320
      Width           =   1452
      _Version        =   1310720
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton btnActualiza 
      Height          =   492
      Left            =   11640
      TabIndex        =   11
      Top             =   1680
      Width           =   1452
      _Version        =   1310720
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Actualiza"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   16
      Picture         =   "frmRH_Vacaciones_Acumuladas.frx":1CFB
      ImageAlignment  =   0
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Actuailización:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   9600
      TabIndex        =   9
      Top             =   1320
      Width           =   1932
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carga Archivo para Actualización de Vacaciones Acumuladas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   15012
   End
End
Attribute VB_Name = "frmRH_Vacaciones_Acumuladas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vPaso As Boolean, strSQL As String, rs As New ADODB.Recordset

Private Sub sbArchivo_Load()
Dim rsExcel As New ADODB.Recordset, pInicializa As Integer, LineasLoad As Long
Dim pNomina As String, pEmpleadoId As String, pDias As Currency

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboNomina.ListCount <= 0 Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

pCliente = cboNomina.ItemData(cboNomina.ListIndex)
pInicializa = 1
LineasLoad = 0
strSQL = "" 'Inicializa Bloque

'Sube: Archivo
Set rsExcel = Excel_Load(txtArchivo.Text, "Vacaciones")
    
Do While Not rsExcel.EOF
  If Not IsNull(rsExcel!Empleado_ID) Then
            pEmpleadoId = Trim(CStr(rsExcel!Empleado_ID))
            pDias = rsExcel!Dias
          
            LineasLoad = LineasLoad + 1
          
             If pInicializa = 1 Then
                pInicializa = 0
                strSQL = strSQL & Space(10) & "exec spRH_Carga_Masiva 'C','RRHH_VACA','" & glogon.Usuario _
                        & "','" & pEmpleadoId & "','" & pNomina & "',1" _
                        & ",'','',''" _
                        & "," & pDias
             Else
                strSQL = strSQL & Space(10) & "exec spRH_Carga_Masiva 'C','RRHH_VACA','" & glogon.Usuario _
                        & "','" & pEmpleadoId & "','" & pNomina & "',0" _
                        & ",'','',''" _
                        & "," & pDias
             End If
             
             If Len(strSQL) > 20000 Then
                Call ConectionExecute(strSQL)
                strSQL = ""
             End If
  
  End If 'Null
  
  rsExcel.MoveNext
Loop
rsExcel.Close

'Procesa Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


'Carga Datos - Revisados
strSQL = "exec spRH_Vacaciones_Cargado_Revisado 'C','RRHH_VACA', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        .Text = rs!Empleado_ID
        .Col = 2
        .Text = rs!Identificacion
        .Col = 3
        .Text = rs!Nombre_Completo
        .Col = 4
        .Text = rs!FECHA_INGRESO
        .Col = 5
        .Text = rs!VACA_ACTUALIZA & ""
        .Col = 6
        .Text = CStr(rs!VACA_ACUMULADAS)
        .Col = 7
        .Text = CStr(rs!Dias)
        .Col = 8
        .Text = CStr(rs!Diferencia)
        rs.MoveNext
    Loop
    rs.Close
End With

Me.MousePointer = vbDefault

If LineasLoad = vGrid.MaxRows Then
    MsgBox "Información Cargada Satisfactoriamente", vbInformation
Else
    MsgBox "Información Cargada Pero con errores en algunas líneas", vbInformation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0

End Sub



Private Sub btnActualiza_Click()
Dim i As Long, pEmpleadoId As String, pDias As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

With vGrid

strSQL = ""

For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    pEmpleadoId = .Text
    .Col = 7
    pDias = .Text

    strSQL = strSQL & Space(10) & "exec spRH_Vacaciones_Ajuste '" & pEmpleadoId & "', " & pDias _
           & ",'" & Format(dtpActualiza.Value, "yyyy-mm-dd") _
           & "', '" & glogon.Usuario & "'"
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
    
Next i

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

End With

Me.MousePointer = vbDefault

MsgBox "Vacaciones acumuladas actualizas Satisfactoriamente!", vbInformation

Call cboNomina_Click

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String

        
Select Case Index
  
  Case 0 'buscar
  
    txtArchivo.Text = ""
    Call sbArchivo_Busca
  
  Case 1 'Cargar
       Call sbArchivo_Load
    
  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: VACACIONES" & vbCrLf _
              & " 3. Columnas.: EMPLEADO_ID, DIAS"
     
     MsgBox vMensaje, vbInformation
         
End Select


End Sub



Private Sub sbArchivo_Busca()


With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        txtArchivo.Text = .FileName
    
End With

End Sub


Private Sub cboNomina_Click()
If vPaso Then Exit Sub

vGrid.MaxRows = 0

txtArchivo.Text = ""

dtpActualiza.Value = fxFechaServidor
dtpActualiza.MaxDate = dtpActualiza.Value
dtpActualiza.MinDate = DateAdd("d", -15, dtpActualiza.Value)

End Sub

Private Sub Form_Load()
vModulo = 23
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


'Nomina

vPaso = True

strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
Call sbCbo_Llena_New(cboNomina, strSQL, False, True)

vPaso = False


Call cboNomina_Click

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




