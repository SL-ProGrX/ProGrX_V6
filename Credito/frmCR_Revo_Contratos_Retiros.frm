VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_Revo_Contratos_Retiros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtDisponible 
      Height          =   435
      Left            =   7560
      TabIndex        =   21
      Top             =   120
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbRetiros 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   12855
      _Version        =   1310723
      _ExtentX        =   22675
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "Datos del Retiro:"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnRetiro 
         Height          =   495
         Left            =   10440
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
         _Version        =   1310723
         _ExtentX        =   3408
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "&Retiro"
         BackColor       =   -2147483633
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
         Picture         =   "frmCR_Revo_Contratos_Retiros.frx":0000
      End
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1560
         Width           =   4215
         _Version        =   1310723
         _ExtentX        =   7435
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   1920
         Width           =   4215
         _Version        =   1310723
         _ExtentX        =   7435
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
         Height          =   315
         Left            =   6960
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtRetiro 
         Height          =   315
         Left            =   6960
         TabIndex        =   19
         Top             =   480
         Width           =   1815
         _Version        =   1310723
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   675
         Left            =   1560
         TabIndex        =   22
         Top             =   840
         Width           =   7215
         _Version        =   1310723
         _ExtentX        =   12726
         _ExtentY        =   1191
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Emitir"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblCuentaTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   600
         TabIndex        =   7
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retiro de ..:"
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
         Index           =   32
         Left            =   5880
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12855
      _Version        =   524288
      _ExtentX        =   22675
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
      MaxCols         =   11
      SpreadDesigner  =   "frmCR_Revo_Contratos_Retiros.frx":0727
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtContrato 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      Top             =   600
      Width           =   5775
      _Version        =   1310723
      _ExtentX        =   10186
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   4080
      TabIndex        =   12
      Top             =   960
      Width           =   5775
      _Version        =   1310723
      _ExtentX        =   10186
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   600
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   960
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDivisa 
      Height          =   315
      Left            =   3120
      TabIndex        =   18
      Top             =   960
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible:"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
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
      Height          =   315
      Left            =   480
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmCR_Revo_Contratos_Retiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean


Private Sub sbContrato_Load()

Me.MousePointer = vbHourglass

On Error GoTo vError

'Contrato
strSQL = "exec spCrd_Revo_Retiros_Consulta " & txtContrato.Text
Call OpenRecordSet(rs, strSQL)

txtCedula.Text = rs!Cedula
txtNombre.Text = rs!Nombre
txtCodigo.Text = rs!Codigo
txtDivisa.Text = rs!cod_Divisa
txtDescripcion.Text = rs!LineaDesc

txtDisponible.Text = Format(rs!Disponible_Real, "Standard")
txtRetiro.Text = Format(0, "Standard")

txtNotas.Text = ""

rs.Close

'Carga Grid

With vGrid
 .MaxRows = 0
 
 strSQL = "exec spCrd_Revo_Retiros_Transito " & txtContrato.Text
 Call OpenRecordSet(rs, strSQL)
 Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    
   .col = 1
   .Text = Format(rs!Registro_Fecha, "yyyy-mm-dd")
   .col = 2
   .Text = Format(rs!Monto, "Standard")
   
   .col = 3
   .Text = rs!Tesoreria_Id & ""
   .col = 4
   .Text = rs!TesoreriaTipo & ""
   .col = 5
   .Text = rs!TesoreriaDocumento & ""
   
   .col = 6
   .Text = rs!BancoDesc & ""
   .col = 7
   .Text = rs!IBAN & ""
   
   .col = 8
   .Text = rs!Registro_Usuario & ""
   .col = 9
   .Text = rs!Autoriza_Usuario & ""
   
   
   .col = 10
   .Text = Format(rs!INT_CORTE, "yyyy-mm-dd")
   .col = 11
   .Text = Format(rs!INT_MONTO, "Standard")
   
   rs.MoveNext
 Loop
 rs.Close
End With

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargaCombos()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

'Consulta todas las cuentas Bancarias
strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
cboTipoDocumento.Text = fxTipoDocumento("TE")

vPaso = False
'
'Call cboTipoDocumento_Click

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnRetiro_Click()

On Error GoTo vError

If CCur(txtRetiro.Text) > CCur(txtDisponible.Text) Then
   MsgBox "El Monto del Retiro es Superior al Disponible!", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass
'spCrd_Revo_Retiros_Movimiento(@Contrato int, @Monto dec(16,2), @BancoId int, @Emite varchar(10), @IBAN varchar(30)
'        ,  @Concepto varchar(100), @Notas varchar(1000), @Usuario varchar(30))
strSQL = "exec spCrd_Revo_Retiros_Movimiento " & txtContrato.Text & ", " & CCur(txtRetiro.Text) & ", " & cboBanco.ItemData(cboBanco.ListIndex) _
       & ", '" & fxTipoDocumento(cboTipoDocumento.Text) & "','" & cboCuenta.ItemData(cboCuenta.ListIndex) _
       & "','RET','" & txtNotas.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Retiro realizado satisfactoriamente!", vbInformation

Call sbContrato_Load

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

Private Sub Form_Load()

vModulo = 1

txtContrato.Text = GLOBALES.gTag

Call sbContrato_Load


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

 Call sbCargaCombos
End Sub


Private Sub txtRetiro_GotFocus()
On Error GoTo vError

txtRetiro.Text = CCur(txtRetiro.Text)

vError:

End Sub

Private Sub txtRetiro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub txtRetiro_LostFocus()
On Error GoTo vError

txtRetiro.Text = Format(CCur(txtRetiro.Text), "Standard")

vError:
End Sub
