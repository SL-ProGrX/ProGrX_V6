VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxC_Cuentas_Correcciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas por Cobrar: Correcciones"
   ClientHeight    =   5835
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9720
      Top             =   240
   End
   Begin XtremeSuiteControls.GroupBox fraOperacion 
      Height          =   2412
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   9612
      _Version        =   1441793
      _ExtentX        =   16954
      _ExtentY        =   4254
      _StockProps     =   79
      Caption         =   "Información General"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin MSComCtl2.FlatScrollBar FlatScrollBarCnt 
         Height          =   252
         Left            =   8520
         TabIndex        =   1
         Top             =   1080
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarPagador 
         Height          =   252
         Left            =   8520
         TabIndex        =   2
         Top             =   1440
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarAutorizado 
         Height          =   252
         Left            =   8520
         TabIndex        =   3
         Top             =   1800
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarConcepto 
         Height          =   252
         Left            =   8520
         TabIndex        =   4
         Top             =   720
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Left            =   3480
         TabIndex        =   16
         Top             =   360
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtConceptoCod 
         Height          =   312
         Left            =   1800
         TabIndex        =   17
         Top             =   720
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtConceptoDesc 
         Height          =   312
         Left            =   3480
         TabIndex        =   18
         Top             =   720
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtContratoCod 
         Height          =   312
         Left            =   1800
         TabIndex        =   19
         Top             =   1080
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtContratoDesc 
         Height          =   312
         Left            =   3480
         TabIndex        =   20
         Top             =   1080
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtPagadorCed 
         Height          =   312
         Left            =   1800
         TabIndex        =   21
         Top             =   1440
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPagadorNom 
         Height          =   312
         Left            =   3480
         TabIndex        =   22
         Top             =   1440
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtAutorizadoCed 
         Height          =   312
         Left            =   1800
         TabIndex        =   23
         Top             =   1800
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAutorizadoNom 
         Height          =   312
         Left            =   3480
         TabIndex        =   24
         Top             =   1800
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8700
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
      Begin VB.Label lblPagador 
         BackStyle       =   0  'Transparent
         Caption         =   "Pagador"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   720
         TabIndex        =   29
         Top             =   1440
         Width           =   1092
      End
      Begin VB.Label lblContrato 
         BackStyle       =   0  'Transparent
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   720
         TabIndex        =   28
         Top             =   1080
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   720
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
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
         Left            =   720
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblAutorizado 
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   720
         TabIndex        =   25
         Top             =   1800
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.GroupBox gbDesembolso 
      Height          =   972
      Left            =   360
      TabIndex        =   5
      Top             =   4680
      Width           =   9732
      _Version        =   1441793
      _ExtentX        =   17166
      _ExtentY        =   1714
      _StockProps     =   79
      Caption         =   "Desembolso y Estado de la Operación"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   312
         Left            =   5880
         TabIndex        =   6
         Top             =   600
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
         Height          =   312
         Left            =   5880
         TabIndex        =   8
         Top             =   240
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta/Desembolso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Emitir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   5040
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   852
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   9732
      _Version        =   1441793
      _ExtentX        =   17166
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Notas del cambio"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnActualiza 
         Height          =   552
         Left            =   8520
         TabIndex        =   14
         ToolTipText     =   "Detalle de Facturas"
         Top             =   240
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   974
         _StockProps     =   79
         Caption         =   "Actualizar"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   552
         Left            =   1080
         TabIndex        =   31
         Top             =   240
         Width           =   7332
         _Version        =   1441793
         _ExtentX        =   12933
         _ExtentY        =   974
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   432
      Left            =   1440
      TabIndex        =   30
      Top             =   240
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   3720
      TabIndex        =   12
      Top             =   240
      Width           =   5532
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
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
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmCxC_Cuentas_Correcciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vEdita              As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso               As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll             As Boolean
Dim mCntPagadorAbierto As Boolean

Private Sub btnActualiza_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If Len(txtNotas.Text) < 10 Then
   MsgBox "Indique una nota válida del cambio a realizar!", vbExclamation
End If

i = MsgBox("Esta seguro de realizar los cambios?", vbYesNo)
If i = vbNo Then Exit Sub

'spCxC_Cuentas_Cambios(@Operacion int, @Usuario varchar(30), @Notas varchar(255)
'                , @ClienteId varchar(30), @PagadorId varchar(30) , @AutorizadorId varchar(30)
'                , @Concepto varchar(10), @Contrato varchar(10)
'                , @BancoId int, @BancoCta varchar(30), @BancoTipo varchar(10))


strSQL = "exec spCxC_Cuentas_Cambios " & txtOperacion.Text & ",'" & glogon.Usuario & "','" & Mid(txtNotas.Text, 1, 255) _
       & "','" & txtCedula.Text & "','" & txtPagadorCed.Text & "','" & txtAutorizadoCed.Text _
       & "','" & txtConceptoCod.Text & "','" & txtContratoCod.Text _
       & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & cboCuenta.ItemData(cboCuenta.ListIndex) _
       & "','" & cboTipoDocumento.ItemData(cboTipoDocumento.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
  If rs!TipoDoc <> "" Then
    Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc, False)
  End If
  MsgBox "Cambios Realizados Satisfactoriamente!", vbInformation
  Unload Me
End If
rs.Close


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

'spSys_Cuentas_Bancarias(@Identificacion varchar(30), @BancoId int, @DivisaCheck smallint = 0)

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCuenta.SetFocus

End Sub

Private Sub FlatScrollBarAutorizado_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarPagador.Tag = "" Then FlatScrollBarPagador.Tag = 0

strSQL = "select Top 1 Per.Cedula,Per.nombre" _
       & " from CxC_Personas Per" _
       & "  inner join CXC_PERSONAS_AUTORIZADOS Pa on Per.Cedula = Pa.Cedula_Autorizado" _
       & " Where Pa.cedula = '" & txtCedula.Text & "'" _

If FlatScrollBarAutorizado.Value > CLng(FlatScrollBarPagador.Tag) Then
   strSQL = strSQL & " and Pa.Cedula_Autorizado > '" & txtAutorizadoCed.Text & "' order by Pa.Cedula_Autorizado asc"
Else
   strSQL = strSQL & " and Pa.Cedula_Autorizado < '" & txtAutorizadoCed.Text & "' order by Pa.Cedula_Autorizado desc"
End If

FlatScrollBarAutorizado.Tag = FlatScrollBarAutorizado.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtAutorizadoCed.Text = rs!Cedula
  txtAutorizadoNom.Text = rs!Nombre
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarCnt_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarCnt.Tag = "" Then FlatScrollBarCnt.Tag = 0

strSQL = "select Top 1 Cnt.Cod_Contrato,'" & txtCedula.Text & "' as 'Cedula'" _
       & " from CxC_Conceptos_Contratos Cnt" _
       & "      inner join CxC_Contratos Cn on Cnt.Cod_Contrato = Cn.cod_Contrato" _
       & "       left join CxC_Personas_Contratos Pc on Cnt.cod_Contrato = Pc.cod_Contrato" _
       & " and Cnt.Cod_Concepto = '" & txtConceptoCod.Text & "' and Pc.Cedula = '" & txtCedula.Text & "'" _
       & " Where Cn.Activo = 1 and Cnt.Cod_Concepto = '" & txtConceptoCod.Text & "'" _
       & "   and (Pc.Cedula is not null or Cn.Suscripcion_Abierta = 1)"

If FlatScrollBarCnt.Value > CLng(FlatScrollBarCnt.Tag) Then
   strSQL = strSQL & " and Cn.cod_contrato > '" & txtContratoCod.Text & "' order by Cn.cod_contrato asc"
Else
   strSQL = strSQL & " and Cn.cod_contrato < '" & txtContratoCod.Text & "' order by Cn.cod_contrato desc"
End If

FlatScrollBarCnt.Tag = FlatScrollBarCnt.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  Call sbContratoDetalle(rs!COD_CONTRATO, rs!Cedula)
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarConcepto_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarConcepto.Tag = "" Then FlatScrollBarConcepto.Tag = 0

strSQL = "Select Top 1 cod_Concepto,Descripcion  from CxC_Conceptos " _
       & " where Activo = 1"

If FlatScrollBarConcepto.Value > CLng(FlatScrollBarConcepto.Tag) Then
   strSQL = strSQL & " and cod_Concepto > '" & txtConceptoCod.Text & "' order by cod_Concepto asc"
Else
   strSQL = strSQL & " and cod_Concepto < '" & txtConceptoCod.Text & "' order by cod_Concepto desc"
End If

FlatScrollBarConcepto.Tag = FlatScrollBarConcepto.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtConceptoCod.Text = rs!cod_Concepto
  txtConceptoDesc.Text = rs!Descripcion
  txtConceptoCod_LostFocus
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarPagador_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarPagador.Tag = "" Then FlatScrollBarPagador.Tag = 0

If mCntPagadorAbierto Then
      strSQL = "select Top 1 Cp.Cedula,Cp.Nombre from CxC_Personas Cp" _
             & " Where Cp.Rol_Pagador = 1"
Else
    strSQL = "select Top 1 Cp.Cedula,Per.nombre" _
           & " from CxC_Contratos_Pagadores Cp inner join  CxC_Contratos Cn on Cp.Cod_Contrato = Cn.Cod_Contrato" _
           & " inner join CxC_Personas Per on Cp.cedula = Per.cedula" _
           & "  left join CxC_Personas_Contratos_Pagadores PcP on Cp.Cod_Contrato = PcP.cod_Contrato" _
           & " and Cp.Cedula = PcP.cedula_Pagador and PcP.cedula = '" & txtCedula.Text & "'" _
           & " Where Cn.Cod_Contrato = '" & txtContratoCod.Text & "'" _
           & " and (PcP.cedula is not null or Cn.Pagadores_Abierto = 1)"
End If

If FlatScrollBarPagador.Value > CLng(FlatScrollBarPagador.Tag) Then
   strSQL = strSQL & " and Cp.Cedula > '" & txtPagadorCed.Text & "' order by Cp.Cedula asc"
Else
   strSQL = strSQL & " and Cp.Cedula < '" & txtPagadorCed.Text & "' order by Cp.Cedula desc"
End If

FlatScrollBarPagador.Tag = FlatScrollBarPagador.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtPagadorCed.Text = rs!Cedula
  txtPagadorNom.Text = rs!Nombre
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbContratoDetalle(pContrato As String, pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

mCntPagadorAbierto = False

strSQL = "select Cnt.Cod_Contrato, Cnt.Descripcion, Cnt.PAGADORES_ABIERTO" _
       & ", isnull(Per.Tasa_Corriente, Cnt.Tasa_Corriente) as 'Tasa_Corriente'" _
       & ", ISNULL(Per.Tasa_Mora,Cnt.Tasa_Mora) as 'Tasa_Mora', isnull(Per.Plazo,Cnt.Plazo) as 'Plazo'" _
       & " from CxC_Contratos Cnt left join CxC_Personas_Contratos Per on  Cnt.Cod_Contrato = Per.cod_contrato" _
       & " and Per.Activo = 1 and Per.Cedula = '" & pCedula & "'" _
       & " Where Cnt.cod_Contrato = '" & pContrato & "'" _
       & "   and (Per.Cedula is not null or Cnt.Suscripcion_Abierta = 1)"
       
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtContratoDesc.Text = rs!Descripcion & ""
   txtContratoCod.Text = rs!COD_CONTRATO & ""
   
   If rs!PAGADORES_ABIERTO = 1 Then
       mCntPagadorAbierto = True
   Else
       mCntPagadorAbierto = False
   End If
   
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Activate()
vModulo = 31
End Sub


Private Sub Form_Load()

vModulo = 31

txtOperacion.Text = GLOBALES.gTag

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub




Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vFecha As Date, iMes As Integer, lngAnio As Long
Dim i As Integer, vTemp As String

On Error GoTo vError

vPaso = True

strSQL = "select * from vCxC_Cuentas_Consulta" _
       & " where Operacion = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 
 vFecha = rs!FechaServer

 mCntPagadorAbierto = IIf((rs!PAGADORES_ABIERTO = 1), True, False)

 
 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 
 lblNombre.Caption = txtNombre.Text
 
 txtConceptoCod.Text = rs!cod_Concepto
 txtConceptoDesc.Text = rs!ConceptoDesc

 txtContratoCod.Text = rs!COD_CONTRATO & ""
 txtContratoDesc.Text = rs!ContratoDesc & ""
 
 txtPagadorCed.Text = rs!cedula_pagador & ""
 txtPagadorNom.Text = rs!PagadorNom & ""
     
 txtAutorizadoCed.Text = rs!Cedula_Autorizado & ""
 txtAutorizadoNom.Text = rs!AutorizadoNom & ""
 
 Call sbCboAsignaDato(cboBanco, rs!BancoDesc, True, rs!Emitir_Banco)
 cboTipoDocumento.Text = fxTipoDocumento(IIf(IsNull(rs!Emitir_Tipo), "OT", rs!Emitir_Tipo))
 
 If Not IsNull(rs!Emitir_Cuenta) Then
    Call sbCboAsignaDato(cboCuenta, rs!CuentaDesc, True, rs!Emitir_Cuenta)
 End If
 
 
 'Bloquea y/o activa campos dependiendo de la configuración del concepto
 Call txtConceptoCod_LostFocus
 

Else
 MsgBox "No existe este número de Operación, verifique!", vbCritical
End If
rs.Close

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbCargaCombos()
Dim strSQL As String

vPaso = True

'Bancos
strSQL = "exec spCxC_Bancos_Autorizados"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
cboTipoDocumento.Text = fxTipoDocumento("TE")

cboCuenta.Clear

vPaso = False

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbCargaCombos
Call sbConsulta

End Sub

Private Sub txtAutorizadoCed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAutorizadoNom.SetFocus


If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "PcP.CEDULA_AUTORIZADO"
   gBusquedas.Orden = "PcP.CEDULA_AUTORIZADO"
   gBusquedas.Consulta = "select PcP.CEDULA_AUTORIZADO,Per.nombre from CXC_PERSONAS_AUTORIZADOS PcP" _
                      & " inner join CxC_Personas Per on PcP.CEDULA_AUTORIZADO = Per.cedula"
   gBusquedas.Filtro = " and PcP.cedula = '" & txtCedula.Text & "'"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtAutorizadoCed.Text = gBusquedas.Resultado
      txtAutorizadoNom.Text = gBusquedas.Resultado2
   End If
End If

End Sub


Private Sub txtAutorizadoNom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And cboBanco.Enabled Then cboBanco.SetFocus
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from CxC_Personas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If


End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCedula.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass


Call cboBanco_Click

strSQL = "select Nombre, Adelanto_Permite,Adelanto_Porcentaje, Adelanto_Modifica, Credito_Limite, Credito_Cerrado" _
       & ", ADELANTO_COMISION, ADELANTO_COMISION_APL" _
       & " from cxc_personas where cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)

txtNombre = rs!Nombre

rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtPagadorCed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtPagadorNom.Enabled Then txtPagadorNom.SetFocus

  If mCntPagadorAbierto Then
        gBusquedas.Columna = "Cedula"
        gBusquedas.Orden = "Cedula"
        gBusquedas.Consulta = "select Cedula,Nombre" _
                           & " from CxC_Personas"
        gBusquedas.Filtro = " and Rol_Pagador = 1"
  Else
        gBusquedas.Columna = "PcP.Cedula_Pagador"
        gBusquedas.Orden = "PcP.Cedula_Pagador"
        gBusquedas.Consulta = "select PcP.Cedula_Pagador,Per.nombre" _
                           & " from CxC_Personas_Contratos_Pagadores PcP" _
                           & " inner join CxC_Personas Per on PcP.cedula_pagador = Per.cedula"
        gBusquedas.Filtro = " and PcP.cod_contrato = '" & txtContratoCod.Text & "' and PcP.cedula = '" & txtCedula.Text & "'"
   End If

   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtPagadorCed.Text = gBusquedas.Resultado
      txtPagadorNom.Text = gBusquedas.Resultado2
   End If

End Sub

Private Sub txtPagadorNom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtAutorizadoCed.Enabled Then txtAutorizadoCed.SetFocus
End Sub



Private Sub txtConceptoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConceptoDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_Concepto"
   gBusquedas.Orden = "cod_Concepto"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select cod_Concepto as 'Concepto',Descripcion  from CxC_Conceptos"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtConceptoCod.Text = gBusquedas.Resultado
      txtConceptoDesc.Text = gBusquedas.Resultado2
      Call txtConceptoCod_LostFocus
   End If
End If

End Sub

Private Sub txtConceptoCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.*,dbo.MyGetdate() as 'FechaServer' " _
       & ", isnull(C.PAGADOR_DEFAULT,'') as 'PagadorId', isnull(P.Nombre,'') as 'PagadorDesc'" _
       & " from CxC_Conceptos C left join CxC_Personas P on C.PAGADOR_DEFAULT = P.cedula" _
       & " where C.cod_Concepto = '" & txtConceptoCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtConceptoDesc.Text = rs!Descripcion
    
   If rs!Requiere_Contrato = 1 Then
      txtContratoCod.Enabled = True
      txtContratoDesc.Enabled = True
   Else
      txtContratoCod.Enabled = False
      txtContratoDesc.Enabled = False
   End If
   
  
   If rs!Proceso_Descuento = 1 Then
      
      txtContratoCod.Enabled = True
      txtContratoDesc.Enabled = True
      
      txtPagadorCed.Enabled = True
      txtPagadorNom.Enabled = True
      
      txtAutorizadoCed.Enabled = True
      txtAutorizadoNom.Enabled = True
      
   Else
      
      txtContratoCod.Enabled = False
      txtContratoDesc.Enabled = False
      
      txtPagadorCed.Enabled = False
      txtPagadorNom.Enabled = False
      
      txtAutorizadoCed.Enabled = False
      txtAutorizadoNom.Enabled = False
   End If
    
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtConceptoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtContratoCod.Enabled Then
    txtContratoCod.SetFocus
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select cod_Concepto,Descripcion  from CxC_Conceptos"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtConceptoCod.Text = gBusquedas.Resultado
      txtConceptoDesc.Text = gBusquedas.Resultado2
      Call txtConceptoCod_LostFocus
   End If
End If
End Sub

Private Sub txtContratoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContratoDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cnt.cod_Contrato"
   gBusquedas.Orden = "Cnt.cod_Contrato"
   gBusquedas.Consulta = "Select Cnt.cod_Contrato,Cnt.Descripcion" _
                       & " from CxC_Personas_Contratos Con inner join CxC_Contratos Cnt on Con.Cod_Contrato = Cnt.cod_contrato"
   gBusquedas.Filtro = " and Con.cedula = '" & txtCedula.Text & "' and Con.cod_contrato in(select cod_contrato" _
                     & " from CxC_Conceptos_Contratos where cod_concepto = '" & txtConceptoCod.Text & "') and Con.Activo = 1"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtContratoCod.Text = gBusquedas.Resultado
      txtContratoDesc.Text = gBusquedas.Resultado2
      
      Call sbContratoDetalle(txtContratoCod.Text, txtCedula.Text)
      
   End If
End If

End Sub


Private Sub txtContratoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And (txtPagadorCed.Enabled) Then txtPagadorCed.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cnt.Descripcion"
   gBusquedas.Orden = "Cnt.Descripcion"
   gBusquedas.Consulta = "Select Cnt.cod_Contrato,Cnt.Descripcion" _
                       & " from CxC_Personas_Contratos Con inner join CxC_Contratos Cnt on Con.Cod_Contrato = Cnt.cod_contrato"
   gBusquedas.Filtro = " and Con.cedula = '" & txtCedula.Text & "' and Con.cod_contrato in(select cod_contrato" _
                     & " from CxC_Conceptos_Contratos where cod_concepto = '" & txtConceptoCod.Text & "') and Con.Activo = 1"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtContratoCod.Text = gBusquedas.Resultado
      txtContratoDesc.Text = gBusquedas.Resultado2
      
      Call sbContratoDetalle(txtContratoCod.Text, txtCedula.Text)
   End If
End If

End Sub

