VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Deducciones_Load 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargado Masivo de Deducciones"
   ClientHeight    =   8220
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.RadioButton rbFormato 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Microsoft Excel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkLimpia 
      Height          =   252
      Left            =   7680
      TabIndex        =   11
      Top             =   2040
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Elimina Información anterior"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9480
      Top             =   0
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   312
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   5532
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   555
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   979
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   0
      Left            =   7680
      TabIndex        =   2
      Top             =   2400
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Deducciones_Load.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   1
      Left            =   8160
      TabIndex        =   3
      Top             =   2400
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Deducciones_Load.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   2
      Left            =   8640
      TabIndex        =   4
      Top             =   2400
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Deducciones_Load.frx":0E19
   End
   Begin XtremeSuiteControls.FlatEdit txtEntidadDesc 
      Height          =   312
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7853
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEntidad 
      Height          =   312
      Left            =   2040
      TabIndex        =   9
      Top             =   2040
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3855
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   12015
      _Version        =   524288
      _ExtentX        =   21193
      _ExtentY        =   6800
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
      MaxCols         =   494
      SpreadDesigner  =   "frmRH_Deducciones_Load.frx":1532
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   492
      Left            =   9360
      TabIndex        =   13
      Top             =   7560
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Deducciones_Load.frx":1BAA
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   492
      Left            =   10680
      TabIndex        =   14
      Top             =   7560
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Deducciones_Load.frx":22D1
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   960
      TabIndex        =   15
      Top             =   7800
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   312
      Left            =   2520
      TabIndex        =   16
      Top             =   7800
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
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
   Begin XtremeSuiteControls.FlatEdit txtErr 
      Height          =   312
      Left            =   3480
      TabIndex        =   17
      Top             =   7800
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
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
   Begin XtremeSuiteControls.ComboBox cboBaseApl 
      Height          =   312
      Left            =   2040
      TabIndex        =   21
      Top             =   1560
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.RadioButton rbFormato 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   24
      Top             =   3000
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Archivo ProGrX RRHH"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Top             =   3000
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Formato"
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
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   0
      Left            =   840
      TabIndex        =   22
      Top             =   1560
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Casos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   2520
      TabIndex        =   20
      Top             =   7560
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   7800
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Errores"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   11
      Left            =   3480
      TabIndex        =   18
      Top             =   7560
      Width           =   972
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Entidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   3
      Left            =   840
      TabIndex        =   10
      Top             =   2040
      Width           =   1092
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
      Left            =   840
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
      Left            =   840
      TabIndex        =   6
      Top             =   2400
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carga Archivo para Deducciones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
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
      TabIndex        =   5
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmRH_Deducciones_Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vProcessId As String
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    vProcessId = ""
End Sub


Private Sub sbCarga_Listado_Excel()
Dim rsExcel As New ADODB.Recordset
Dim pInicial As Boolean

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "" 'Inicializa Bloque

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")

pInicial = True

vProcessId = "E-" & txtEntidad.Text & "_" & cboNomina.ItemData(cboNomina.ListIndex)
vProcessId = Mid(vProcessId, 1, 10)


strSQL = ""
Do While Not rsExcel.EOF

  If Not IsNull(rsExcel!IDENTIFICACION) Then
  
        If pInicial Then
          pInicial = False
          strSQL = "exec spRRHH_Carga_Masiva 'R','" & vProcessId & "','" & glogon.Usuario _
              & "','" & rsExcel!IDENTIFICACION & "','','" & Trim(rsExcel!Tipo) & "','" & rsExcel!Codigo _
              & "','Pending'," & rsExcel!Monto _
              & ",'" & Format(rsExcel!Inicio, "yyyy/mm/dd") & "','" & Format(rsExcel!Corte, "yyyy/mm/dd") _
              & "',1"
        Else
          strSQL = strSQL & Space(10) & "exec spRRHH_Carga_Masiva 'R','" & vProcessId & "','" & glogon.Usuario _
              & "','" & rsExcel!IDENTIFICACION & "','','" & Trim(rsExcel!Tipo) & "','" & rsExcel!Codigo _
              & "','Pending'," & rsExcel!Monto _
              & ",'" & Format(rsExcel!Inicio, "yyyy/mm/dd") & "','" & Format(rsExcel!Corte, "yyyy/mm/dd") _
              & "',0"
        End If
  
  End If
  
  If Len(strSQL) > 20000 Then
     Call ConectionExecute(strSQL)
     strSQL = ""
  End If

  rsExcel.MoveNext
Loop

rsExcel.Close
            
'Procesa el Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If
            
            
'Vincula y Regresa Set
Call sbCarga_Listado_Revisa

Me.MousePointer = vbDefault


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0

End Sub



Private Sub sbCarga_Listado_ProGrX_RRHH()
Dim fn, strCadena As String
Dim pInicial As Boolean
Dim pIdentificacion As String, pTipo As String, pInicio As Date, pCorte As Date
Dim pCodigo As String, pMonto As Currency

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "" 'Inicializa Bloque

pInicial = True

vProcessId = "E-" & txtEntidad.Text & "_" & cboNomina.ItemData(cboNomina.ListIndex)
vProcessId = Mid(vProcessId, 1, 10)


fn = FreeFile

Open frmContenedor.CD.FileName For Input As #fn   ' Lee el archivo.
 
strSQL = ""

Do While Not EOF(fn)
   Input #fn, strCadena
 
    pIdentificacion = Trim(Mid(strCadena, 1, 20))
    pCodigo = Trim(Mid(strCadena, 21, 10))
    pTipo = Mid(strCadena, 31, 1)
    pMonto = CCur(Mid(strCadena, 33, 10)) / 100
    pInicio = Mid(strCadena, 44, 10)
    pCorte = Mid(strCadena, 55, 10)
 
 
 
  If pIdentificacion <> "" Then
  
        If pInicial Then
          pInicial = False
          strSQL = "exec spRRHH_Carga_Masiva 'R','" & vProcessId & "','" & glogon.Usuario _
              & "','" & pIdentificacion & "','" & pCodigo & "','" & pTipo & "','" & pCodigo _
              & "','Pending'," & pMonto _
              & ",'" & Format(pInicio, "yyyy/mm/dd") & "','" & Format(pCorte, "yyyy/mm/dd") _
              & "',1"
        Else
          strSQL = strSQL & Space(10) & "exec spRRHH_Carga_Masiva 'R','" & vProcessId & "','" & glogon.Usuario _
              & "','" & pIdentificacion & "','" & pCodigo & "','" & pTipo & "','" & pCodigo _
              & "','Pending'," & pMonto _
              & ",'" & Format(pInicio, "yyyy/mm/dd") & "','" & Format(pCorte, "yyyy/mm/dd") _
              & "',0"
        End If
  
  End If
  
  If Len(strSQL) > 20000 Then
     Call ConectionExecute(strSQL)
     strSQL = ""
  End If
 
 
 Loop
Close #fn


'Procesa el Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If
            
            
'Vincula y Regresa Set
Call sbCarga_Listado_Revisa

Me.MousePointer = vbDefault


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0

End Sub




Private Sub sbCarga_Listado_Revisa()
Dim curTotal As Currency, pErrores As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

            
'Vincula y Regresa Set
            
strSQL = "exec spRRHH_Carga_Masiva_Revisa 'R','" & vProcessId & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

curTotal = 0
pErrores = 0

With vGrid
    .MaxRows = 0

     Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
   
        .Col = 1 'Id
        .Text = rs!Llave_01
        .Col = 2 'Nombre/Error
        .Text = rs!Ref_03
        
        .Col = 3 'Tipo
        .Text = rs!Ref_01
        .Col = 4 'Valor
        .Text = Format(rs!Monto_01, "Standard")
        .Col = 5 'Codigo
        .Text = rs!Llave_02
        
        .Col = 6 'Inicio
        .Text = Format(rs!Fecha_01, "yyyy/MM/dd")
        .Col = 7 'Corte
        .Text = Format(rs!Fecha_02, "yyyy/MM/dd")
      
      .Col = 2
      If Mid(.Text, 1, 6) = "Error!" Then
        pErrores = pErrores + 1
      Else
        curTotal = curTotal + rs!Monto_01
      End If
      
      
      rs.MoveNext
    Loop
    rs.Close
End With

Me.MousePointer = vbDefault

txtMonto.Text = Format(curTotal, "Standard")
txtCasos.Text = Format(vGrid.MaxRows, "###,###,##0")
txtErr.Text = Format(pErrores, "###,###,##0")

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0

End Sub



Private Sub sbProcesar()
Dim lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Carga_Masiva_Procesa 'R','" & vProcessId & "','" & glogon.Usuario _
       & "'," & chkLimpia.Value & ",'" & Mid(cboBaseApl.Text, 1, 1) & "'"

Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Carga Masiva Deducciones " & cboNomina.ItemData(cboNomina.ListIndex) & ", Entidad: " & txtEntidad.Text & ", Líneas(" & vGrid.MaxRows & ")")

Me.MousePointer = vbDefault

MsgBox "Listado de Deducciones Procesado Satisfactoriamente...", vbInformation

txtArchivo.Text = ""
vGrid.MaxRows = 0

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen deducciones cargadas...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
End Sub

Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String

        
Select Case Index
  
  Case 0 'buscar
  
    txtArchivo.Text = ""
    Call sbBuscaArchivo(1)
  
  Case 1 'Cargar
  
    If rbFormato(0).Value Then
       Call sbCarga_Listado_Excel
    Else
       Call sbCarga_Listado_ProGrX_RRHH
    End If
    
  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: IMPORT" & vbCrLf _
              & " 3. Columnas.: IDENTIFICACION, CODIGO, TIPO, MONTO, INICIO, CORTE"
     
     MsgBox vMensaje, vbInformation
         
End Select


End Sub



Private Sub sbBuscaArchivo(vTipo As Integer)


With frmContenedor.CD
    If rbFormato(0).Value Then
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
    
    Else
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Deducciones [Texto]..."
        .Filter = "*.txt"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) <> "TXT" Then
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If

        txtArchivo.Text = .FileName

End If
End With

End Sub


Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub


Private Sub cboNomina_Click()
Call sbLimpia
End Sub

Private Sub Form_Load()
vModulo = 23

 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
vGrid.MaxRows = 0
vGrid.MaxCols = 7

cboBaseApl.Clear
cboBaseApl.AddItem "Nómina"
cboBaseApl.AddItem "Mensual"
cboBaseApl.Text = "Mensual"

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub sbInicializa()

On Error GoTo vError


'Nomina
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "COD_ER"
    gBusquedas.Orden = "COD_ER"
    gBusquedas.Consulta = "select COD_ER, NOMBRE FROM RH_ENTIDADES_RELACIONADAS"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    txtEntidad.Text = gBusquedas.Resultado
    txtEntidadDesc.Text = gBusquedas.Resultado2
End If
End Sub
