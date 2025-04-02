VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCR_AbonoMasivo_Manual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicación Masiva de Abonos con Archivo"
   ClientHeight    =   7656
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11352
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7656
   ScaleWidth      =   11352
   Begin XtremeSuiteControls.RadioButton rbOpcion 
      Height          =   252
      Index           =   0
      Left            =   2520
      TabIndex        =   25
      Top             =   1320
      Width           =   2292
      _Version        =   1245185
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Rebaja de un Fondo?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkExcel 
      Height          =   252
      Left            =   9480
      TabIndex        =   0
      Top             =   1320
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Formato Excel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   11172
      _Version        =   524288
      _ExtentX        =   19706
      _ExtentY        =   5948
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
      MaxCols         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_AbonoMasivo_Manual.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   6852
      _Version        =   1245185
      _ExtentX        =   12086
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   312
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   6852
      _Version        =   1245185
      _ExtentX        =   12086
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   432
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   6852
      _Version        =   1245185
      _ExtentX        =   12086
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   960
      TabIndex        =   5
      Top             =   7200
      Width           =   1572
      _Version        =   1245185
      _ExtentX        =   2773
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   312
      Left            =   2520
      TabIndex        =   6
      Top             =   7200
      Width           =   972
      _Version        =   1245185
      _ExtentX        =   1714
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSocios 
      Height          =   312
      Left            =   3480
      TabIndex        =   7
      Top             =   7200
      Width           =   972
      _Version        =   1245185
      _ExtentX        =   1714
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtContratos 
      Height          =   312
      Left            =   4440
      TabIndex        =   8
      Top             =   7200
      Width           =   972
      _Version        =   1245185
      _ExtentX        =   1714
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   312
      Left            =   2520
      TabIndex        =   9
      Top             =   2880
      Width           =   1812
      _Version        =   1245185
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
      Height          =   312
      Left            =   4320
      TabIndex        =   10
      Top             =   2880
      Width           =   5052
      _Version        =   1245185
      _ExtentX        =   8911
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   0
      Left            =   9480
      TabIndex        =   11
      Top             =   1680
      Width           =   492
      _Version        =   1245185
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Appearance      =   16
      Picture         =   "frmCR_AbonoMasivo_Manual.frx":069F
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   1
      Left            =   9960
      TabIndex        =   12
      Top             =   1680
      Width           =   492
      _Version        =   1245185
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Appearance      =   16
      Picture         =   "frmCR_AbonoMasivo_Manual.frx":10BD
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   2
      Left            =   10440
      TabIndex        =   13
      Top             =   1680
      Width           =   492
      _Version        =   1245185
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Appearance      =   16
      Picture         =   "frmCR_AbonoMasivo_Manual.frx":1A80
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   492
      Left            =   7800
      TabIndex        =   14
      Top             =   6960
      Width           =   1572
      _Version        =   1245185
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_AbonoMasivo_Manual.frx":225F
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   492
      Left            =   9360
      TabIndex        =   15
      Top             =   6960
      Width           =   1572
      _Version        =   1245185
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_AbonoMasivo_Manual.frx":2A37
   End
   Begin XtremeSuiteControls.CheckBox chkFondoGeneral 
      Height          =   492
      Left            =   9480
      TabIndex        =   24
      Top             =   2400
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplica como Fondo General"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.RadioButton rbOpcion 
      Height          =   252
      Index           =   1
      Left            =   5760
      TabIndex        =   26
      Top             =   1320
      Width           =   2292
      _Version        =   1245185
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Con Cuenta Contable"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   732
      Left            =   2640
      TabIndex        =   27
      Top             =   240
      Width           =   8532
      _Version        =   1245185
      _ExtentX        =   15049
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Aplicación Masiva de Abonos versus Plan o Cuenta Contable"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
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
      Left            =   1200
      TabIndex        =   23
      Top             =   2160
      Width           =   1332
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
      Left            =   1200
      TabIndex        =   22
      Top             =   1680
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
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
      Index           =   3
      Left            =   1200
      TabIndex        =   21
      Top             =   2520
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   20
      Top             =   2880
      Width           =   1215
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
      Top             =   7200
      Width           =   852
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
      TabIndex        =   18
      Top             =   6960
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Existe ?"
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
      Index           =   2
      Left            =   3480
      TabIndex        =   17
      Top             =   6960
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ident.?"
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
      Index           =   3
      Left            =   4440
      TabIndex        =   16
      Top             =   6960
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11532
   End
End
Attribute VB_Name = "frmCR_AbonoMasivo_Manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFecha As Date, vPaso As Boolean, mContrato As Long


Private Sub sbLimpia()
On Error GoTo vError

    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtSocios.Text = 0
    txtContratos.Text = 0
    txtArchivo.Text = ""

   
vError:
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
    If chkExcel.Value = vbChecked Then
       Call sbBuscaArchivo(1)
    End If
  
  Case 1 'Cargar
    If chkExcel.Value = vbChecked Then
       Call sbCargaDeducciones(1)
    End If
    
  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: Import" & vbCrLf _
              & " 3. Columnas.: OPERACION, ABONO"
     
     MsgBox vMensaje, vbInformation
         
End Select


End Sub


Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub

Private Sub cboOperadora_Click()
Dim strSQL As String

If vPaso Or cboOperadora.ListCount = 0 Then Exit Sub

strSQL = "select rtrim(cod_plan) as 'IdX', rtrim(descripcion) + space(10) + '[' + rtrim(cod_Plan) + ']' as 'ItmX'" _
       & " from fnd_planes where deduce_independiente = 1 and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
vPaso = True

Call sbCbo_Llena_New(cboPlan, strSQL, False, True)

vPaso = False

End Sub


Private Sub chkExcel_Click()
 Call sbLimpia
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer
Dim vProceso As Long

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

mFecha = fxFechaServidor

vPaso = True
    
    strSQL = "select cod_operadora as IdX, descripcion as ItmX from FND_Operadoras"
    Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

txtArchivo.Text = ""

vGrid.MaxCols = 5
vGrid.MaxRows = 0

vPaso = False


Call cboOperadora_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbCargaDeducciones(vTipo As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

Dim pCedula As String, pNombre As String, pAbono As Currency, pOperacion As Long
Dim pOperadora As Integer, pPlan As String, pLinea As Long

Dim strCadena As String, curMonto As Currency
Dim lCasos As Long
Dim strMonto  As String
Dim strCedula As String
Dim strNombre As String
Dim i As Integer, vCampos As Boolean



On Error GoTo vError


vGrid.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboOperadora.ListCount <= 0 Then
    MsgBox "No existe ninguna Operadora, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If

If cboPlan.ListCount <= 0 Then
   MsgBox "No existe ningun plan, no se puede procesar el archivo...", vbCritical
   Exit Sub
End If


Me.MousePointer = vbHourglass


pOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
pPlan = cboPlan.ItemData(cboPlan.ListIndex)


txtContratos.Text = 0
txtSocios.Text = 0
txtMonto.Text = 0
txtCasos.Text = 0

curMonto = 0
lCasos = 0 'Total

If vTipo = 1 Then 'Archivo de excel

        Set rs = Excel_Load(txtArchivo.Text, "Import")
            
        'Validaciónn del Archivo
        vCampos = False
        For i = 0 To rs.Fields.Count
             
            If UCase(LCase(rs.Fields(i).Name)) = "OPERACION" Then
               vCampos = True
            End If
             
             If vCampos Then Exit For
        Next i
        
        If Not vCampos Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Operacion, Abono ¦ Nombre de la Hoja = Import"
           Exit Sub
        End If
        
        vCampos = False
        For i = 0 To rs.Fields.Count
             
            If UCase(LCase(rs.Fields(i).Name)) = "ABONO" Then
               vCampos = True
            End If
             
             If vCampos Then Exit For
        Next i
        
        If Not vCampos Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Operacion, Abono ¦ Nombre de la Hoja = Import"
           Exit Sub
        End If
        
        'FIN: Validación del Archivo
        
        'Sube, Revisa y Carga
        With vGrid
            
            pLinea = 0
            strSQL = ""
            
            Do While Not rs.EOF
              If IsNumeric(rs!Operacion & "") Then
                pOperacion = rs!Operacion
                pAbono = rs!Abono

                pLinea = pLinea + 1
                
                If pLinea = 1 Then
                    strSQL = strSQL & Space(10) & "exec spSys_Carga_Masiva_Sube 'C','CrdAplAbo','" & pOperacion & "','','" & glogon.Usuario _
                           & "','','',''," _
                           & pAbono & ",0,0,''," _
                           & "Null, Null, Null,1"
                Else
                    strSQL = strSQL & Space(10) & "exec spSys_Carga_Masiva_Sube 'C','CrdAplAbo','" & pOperacion & "','','" & glogon.Usuario _
                           & "','','',''," _
                           & pAbono & ",0,0,''," _
                           & "Null, Null, Null,0"
                End If
                
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
        
        'Revisa Lote y lo Carga
        strSQL = "exec spSys_Carga_Masiva_Revisa_AplMasCrd_Abono 'C','CrdAplAbo','','','" & glogon.Usuario & "'"
        Call OpenRecordSet(rs, strSQL)
        If glogon.error Then
           Exit Sub
        End If

            Do While Not rs.EOF
              
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .col = 1
                    .Text = CStr(rs!id_solicitud)
                    .col = 2
                    .Text = rs!Codigo
                    .col = 3
                    .Text = rs!Cedula
                    .col = 4
                    .Text = rs!Nombre
                    .col = 5
                    .Text = Format(rs!Abono, "Standard")
                    
                    curMonto = curMonto + rs!Abono
                    txtMonto.Text = Format(curMonto, "Standard")
                    txtCasos.Text = txtCasos.Text + 1
              
              rs.MoveNext
            Loop
            rs.Close
        
        
    End With 'vGrid


        
End If 'end if tipo archivo


'Totales
txtMonto.Text = Format(curMonto, "Standard")
Me.MousePointer = vbDefault

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub

Private Sub sbProcesar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDoc As String, vNumDoc As String, vTipo As String
Dim vCuenta  As String, vOperadora As Long, vPlan As String

On Error GoTo vError

If vGrid.MaxRows = 0 Then Exit Sub

Me.MousePointer = vbHourglass

vPlan = cboPlan.ItemData(cboPlan.ListIndex)
vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)
 
Select Case True
Case rbOpcion.Item(0).Value
    vTipo = "P"
Case rbOpcion.Item(0).Value
    vTipo = "C"
End Select
 
 
strSQL = "exec spSys_Carga_Masiva_Aplica_AplMasCrd_Abono 'C','CrdAplAbo','','','" & glogon.Usuario & "'," _
                            & vOperadora & ",'" & vPlan & "','" & vCuenta & "'," & chkFondoGeneral.Value _
                            & ",'" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)
  vTipoDoc = rs!Tipo_Documento
  vNumDoc = rs!Num_Documento

rs.Close
 
Me.MousePointer = vbDefault
MsgBox "Proceso Aplicado Satisfactoriamente... Registros Procesados :" & vGrid.MaxRows

Call sbLimpia
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub

Private Sub sbBuscaArchivo(vTipo As Integer)


With frmContenedor.CD
    If vTipo = 1 Or chkExcel.Value = vbChecked Then
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
        
        If UCase(Right(.FileName, 3)) = "XLS" Then
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        'If UCase(Right(.FileName, 3)) <> "TXT" Or UCase(Right(.FileName, 3)) <> "DAT" Then
         '   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
         '   Exit Sub
        'End If

        txtArchivo.Text = .FileName

End If
End With

End Sub






Private Function fxExisteContrato(vCedula As String) As Boolean

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_contrato from fnd_contratos where cedula = '" & vCedula _
         & "' And cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) & "" _
         & " and cod_plan = '" & SIFGlobal.fxCodText(cboPlan.Text) & "' and estado ='A'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF Then
    fxExisteContrato = False
Else
    fxExisteContrato = True
    mContrato = rs!COD_CONTRATO
End If
rs.Close
End Function


Public Sub sbBitacoraPlanilla(pTransaccion As String, pInstitucion As Long, pProceso As Long _
                , pGestion As String, pMonto As Currency, pPlan As String, Optional pDocumento As String = "")
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select isnull(max(id_seq),0) + 1 as Consecutivo from fnd_prm_bitacora" _
       & " where cod_institucion = " & pInstitucion & " and cod_plan  = '" & pPlan & "' and proceso = " & pProceso


Call OpenRecordSet(rs, strSQL)
    strSQL = "insert fnd_prm_bitacora(id_seq,cod_institucion,proceso,cod_plan,gestion,transaccion,documento,usuario,fecha,casos,monto) values(" _
           & rs!Consecutivo & "," & pInstitucion & "," & pProceso & ",'" & pPlan & "','" & pGestion & "','" & pTransaccion _
           & "','" & pDocumento & "','" & glogon.Usuario & "',dbo.MyGetdate(),'" & txtCasos.Text & "'," & CCur(pMonto) & ")"
rs.Close


Call ConectionExecute(strSQL)

End Sub



Private Sub sbFndAsiento(vInstitucion As Long, vProceso As Long, vOperadora As Long, vPlan As String _
        , vCuentaPlanilla As String, Optional vComprobante As String = "")
Dim strSQL As String '


strSQL = "exec spFndPlanillaDirectaAsiento " & vProceso & "," & vInstitucion & "," & vOperadora & ",'" & vPlan _
       & "','" & Trim(vCuentaPlanilla) & "','" & vComprobante & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

End Sub


Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   If Len(gCuenta) > 0 Then
      txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
      txtCuenta.Text = fxgCntCuentaFormato(True, gCuenta, 0)
   End If

End If
End Sub


