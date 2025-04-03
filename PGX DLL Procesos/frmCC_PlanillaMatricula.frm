VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmCC_PlanillaMatricula 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Deducciones - Mantenimiento de Matriculas"
   ClientHeight    =   8004
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   13416
   LinkTopic       =   "Form1"
   ScaleHeight     =   8004
   ScaleWidth      =   13416
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   1212
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   13212
      _Version        =   1310720
      _ExtentX        =   23304
      _ExtentY        =   2138
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Buscar"
      Item(0).ControlCount=   13
      Item(0).Control(0)=   "btnTool(0)"
      Item(0).Control(1)=   "txtProceso"
      Item(0).Control(2)=   "txtCedula"
      Item(0).Control(3)=   "txtDocRef"
      Item(0).Control(4)=   "txtOperacion"
      Item(0).Control(5)=   "txtCodigo"
      Item(0).Control(6)=   "Label2(6)"
      Item(0).Control(7)=   "Label2(5)"
      Item(0).Control(8)=   "Label2(4)"
      Item(0).Control(9)=   "Label2(3)"
      Item(0).Control(10)=   "Label2(2)"
      Item(0).Control(11)=   "Label2(1)"
      Item(0).Control(12)=   "txtNombre"
      Item(1).Caption =   "Bloqueo Masivo"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "Label2(7)"
      Item(1).Control(1)=   "txtArchivo"
      Item(1).Control(2)=   "btnTool(1)"
      Item(1).Control(3)=   "btnTool(2)"
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   312
         Left            =   -69640
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   9252
         _Version        =   1310720
         _ExtentX        =   16319
         _ExtentY        =   550
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnTool 
         Height          =   312
         Index           =   0
         Left            =   12000
         TabIndex        =   6
         Top             =   360
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit txtProceso 
         Height          =   312
         Left            =   4440
         TabIndex        =   7
         ToolTipText     =   "Año + Mes: 202108"
         Top             =   720
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   5880
         TabIndex        =   8
         Top             =   720
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocRef 
         Height          =   312
         Left            =   3000
         TabIndex        =   9
         Top             =   720
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   312
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   7320
         TabIndex        =   18
         Top             =   720
         Width           =   5772
         _Version        =   1310720
         _ExtentX        =   10181
         _ExtentY        =   550
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnTool 
         Height          =   312
         Index           =   1
         Left            =   -60400
         TabIndex        =   21
         ToolTipText     =   "Los campos son cedula, numerooperacion, codigodeduccion, codigoacreedor ¦ Nombre de la Hoja = IMPORT"
         Top             =   720
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnTool 
         Height          =   312
         Index           =   2
         Left            =   -59200
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1310720
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Bloquear Casos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   7
         Left            =   -69640
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Archivo:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   2
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Operación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   3
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Doc.Referencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   4
         Left            =   5880
         TabIndex        =   14
         Top             =   360
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cédula"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   5
         Left            =   7320
         TabIndex        =   13
         Top             =   360
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Nombre"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   372
         Index           =   6
         Left            =   4440
         TabIndex        =   12
         Top             =   360
         Width           =   1092
         _Version        =   1310720
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Proceso"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.CheckBox chkActivos 
      Height          =   252
      Left            =   7680
      TabIndex        =   4
      Top             =   1320
      Width           =   2172
      _Version        =   1310720
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Casos Activos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Value           =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   10800
      Top             =   240
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   12372
      _Version        =   524288
      _ExtentX        =   21823
      _ExtentY        =   8700
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
      MaxCols         =   14
      SpreadDesigner  =   "frmCC_PlanillaMatricula.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   5772
      _Version        =   1310720
      _ExtentX        =   10181
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   312
      Index           =   3
      Left            =   10320
      TabIndex        =   23
      Top             =   1320
      Width           =   2892
      _Version        =   1310720
      _ExtentX        =   5101
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Crear Archivo de Matricula Total"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1092
      _Version        =   1310720
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Institución"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento de Matriculas"
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
      Height          =   372
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   9732
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12612
   End
End
Attribute VB_Name = "frmCC_PlanillaMatricula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean


Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

strSQL = "SELECT ID_REFERENCIA, TIPO, FECHA_PROCESO, COD_DEDUCCION, OPERACION" _
       & " , FORMAT(FORMALIZACION, 'yyyy-MM-dd') , MONTO, CUOTA, PLAZO, TASA, NREFENCIA_EXT" _
       & " , CEDULA, NOMBRE, ISNULL(B_INDICA,0)" _
       & " From vPrm_Matricula" _
       & " Where COD_INSTITUCION = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " AND ISNULL(B_INDICA,0) = " & IIf(chkActivos.Value = xtpChecked, 0, 1) _

If Len(Trim(txtCedula.Text)) > 0 And fxSIFValidaCadena(txtCedula.Text) Then
   strSQL = strSQL & " AND CEDULA LIKE '%" & txtCedula.Text & "%'"
End If
       
If Len(Trim(txtNombre.Text)) > 0 And fxSIFValidaCadena(txtNombre.Text) Then
   strSQL = strSQL & " AND NOMBRE LIKE '%" & txtNombre.Text & "%'"
End If
       
If Len(Trim(txtCodigo.Text)) > 0 And fxSIFValidaCadena(txtCodigo.Text) Then
   strSQL = strSQL & " AND COD_DEDUCCION LIKE '%" & txtCodigo.Text & "%'"
End If
       
If Len(Trim(txtOperacion.Text)) > 0 And fxSIFValidaCadena(txtOperacion.Text) Then
   strSQL = strSQL & " AND OPERACION LIKE '%" & txtOperacion.Text & "%'"
End If
       
If Len(Trim(txtDocRef.Text)) > 0 And fxSIFValidaCadena(txtDocRef.Text) Then
   strSQL = strSQL & " AND NREFENCIA_EXT LIKE '%" & txtDocRef.Text & "%'"
End If
       
       
If IsNumeric(txtProceso.Text) Then
   strSQL = strSQL & " AND FECHA_PROCESO = " & txtProceso.Text
End If

vPaso = True

Call sbCargaGrid(vGrid, 14, strSQL)

vPaso = False

If vGrid.MaxRows > 0 Then
   vGrid.MaxRows = vGrid.MaxRows - 1
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbBloqueoMasivo()
Dim pCedula As String, pOperacion As String, pCodigoDeduc As String
Dim vCampos As Boolean, i As Long


On Error GoTo vError

Me.MousePointer = vbHourglass


Set rs = Excel_Load(txtArchivo.Text, "IMPORT")
    
'Validaciónn del Archivo
'cedula  numerooperacion codigodeduccion cuota   quincena    codigoacreedor


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "CEDULA" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos son cedula, numerooperacion, codigodeduccion, codigoacreedor ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = UCase("numerooperacion") Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos son cedula, numerooperacion, codigodeduccion, codigoacreedor ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = UCase("codigodeduccion") Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos son cedula, numerooperacion, codigodeduccion, codigoacreedor ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


'FIN: Validación del Archivo




'Sube, Revisa y Carga
strSQL = ""
i = 0
    
Do While Not rs.EOF
If Trim(rs!Cedula) <> "" Then
  pCedula = rs!Cedula
  pOperacion = rs!numerooperacion
  pCodigoDeduc = rs!codigodeduccion
      
  i = i + 1
  
  strSQL = strSQL & Space(10) & " exec spPrm_Matricula_Bloqueo_Masivo " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
        & ",'" & pCedula & "','" & pCodigoDeduc & "','" & pOperacion & "','" & glogon.Usuario & "'"
  
  If Len(strSQL) > 20000 Then
     Call ConectionExecute(strSQL)
     If glogon.error Then
        Exit Sub
     End If
     strSQL = ""
  End If
  
End If
rs.MoveNext
Loop
rs.Close

'Procesa Ultimo Bloque

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   If glogon.error Then
      Exit Sub
   End If
   strSQL = ""
End If


txtArchivo.Text = ""
tcMain.Item(0).Selected = True

Me.MousePointer = vbDefault

MsgBox "Bloqueo Masivo realizado Satisfactoriamente, Casos Bloqueados: " & i, vbInformation

Call sbBuscar

Exit Sub

vError:
    txtArchivo.Text = ""
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbMatriculaTotal_Archivo()

Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

'********************************************
'* Formato INTEGRA Tesoreria Nacional       *
'********************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor



vArchivo = ""

Me.MousePointer = vbHourglass

GLOBALES.gInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)


strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & cboInstitucion.Text

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & cboInstitucion.Text
On Error GoTo vError


  
vArchivo = "MD-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-01_TOTAL.csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


strSQL = "exec spPrm_Formato_Integra_New_Matricula_Total " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 Print #fnFile, vCadena
 
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub btnTool_Click(Index As Integer)

Select Case Index
    Case 0
        Call sbBuscar
    Case 1 'Buscar Archivo
    
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
    
    Case 2 'Bloqueo Masivo
        If txtArchivo.Text <> "" Then
            Call sbBloqueoMasivo
        End If
    Case 3 'Archivo Matricula Completa
        Call sbMatriculaTotal_Archivo
End Select

End Sub

Private Sub Form_Load()
 
 vModulo = 10
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

tcMain.Item(0).Selected = True

 Call Formularios(Me)
 Call RefrescaTags(Me)

 vGrid.MaxCols = 14
 
End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 550
vGrid.Height = Me.Height - (vGrid.Top + 850)
tcMain.Width = vGrid.Width

End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0

strSQL = "select COD_INSTITUCION AS IDX, DESCRIPCION AS ITMX" _
       & " FROM INSTITUCIONES WHERE ACTIVA = 1" _
       & " ORDER BY DESCRIPCION"
Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

Call sbBuscar
End Sub



Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call sbBuscar
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call sbBuscar
End If
End Sub


Private Sub txtDocRef_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call sbBuscar
End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call sbBuscar
End If
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call sbBuscar
End If

End Sub


Private Sub txtProceso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call sbBuscar
End If
End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If Col = 14 Then
   vGrid.Row = Row
   vGrid.Col = 14
   If vGrid.Value = vbChecked Then
      vGrid.Col = 1
      strSQL = "exec spPrm_Matricula_Bloqueo " & vGrid.Text & ",'" & glogon.Usuario & "'"
      Call ConectionExecute(strSQL)
      
      MsgBox "Caso Bloqueado! Referencia Id: " & vGrid.Text
      
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

   End If

End If

End Sub
