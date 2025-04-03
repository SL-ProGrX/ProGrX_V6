VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmAH_RegistraAhorro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registra Aportes a una persona"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9735
   Icon            =   "frmAH_RegistraAhorros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnRefresh 
      Height          =   375
      Left            =   8280
      TabIndex        =   22
      ToolTipText     =   "Revisa si fue autorizada!"
      Top             =   3585
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_RegistraAhorros.frx":030A
   End
   Begin XtremeSuiteControls.CheckBox chkAjuste 
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   960
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Es un Ajuste?"
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
      Alignment       =   1
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9000
      Top             =   240
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   9495
      _Version        =   1572864
      _ExtentX        =   16743
      _ExtentY        =   2773
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   2772
         _Version        =   1572864
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.FlatEdit txtTotalCajas 
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   1815
         _Version        =   1572864
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   5412
         _Version        =   1572864
         _ExtentX        =   9546
         _ExtentY        =   1397
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
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   0
         Left            =   6720
         TabIndex        =   4
         Top             =   480
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Pago"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_RegistraAhorros.frx":0A0A
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   1
         Left            =   7680
         TabIndex        =   5
         Top             =   480
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Aplicar"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_RegistraAhorros.frx":0EB7
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   2
         Left            =   8520
         TabIndex        =   6
         Top             =   480
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Cancelar"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_RegistraAhorros.frx":168F
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.CheckBox chkReciboDigital 
         Height          =   255
         Left            =   7080
         TabIndex        =   29
         ToolTipText     =   "Enviar Recibo Digital"
         Top             =   120
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Recibo Digital?"
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
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total ..:"
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
         Index           =   4
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas ..:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento ..:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1452
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   4920
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   4920
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3360
      TabIndex        =   14
      Top             =   240
      Width           =   5292
      _Version        =   1572864
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1680
      TabIndex        =   15
      Top             =   240
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtDivisa 
      Height          =   315
      Left            =   6720
      TabIndex        =   17
      Top             =   2640
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAporteAutorizado 
      Height          =   315
      Left            =   4920
      TabIndex        =   20
      Top             =   2160
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      ToolTipText     =   "Revisa si fue autorizada!"
      Top             =   3120
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Solicita Autorización"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtGestionId 
      Height          =   315
      Left            =   4920
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
      _Version        =   1572864
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtGestionEstado 
      Height          =   315
      Left            =   6720
      TabIndex        =   25
      Top             =   3600
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   330
      Left            =   4920
      TabIndex        =   27
      Top             =   1680
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Proceso..:"
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
      Index           =   7
      Left            =   2400
      TabIndex        =   28
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Gestión ..:"
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
      Index           =   6
      Left            =   2280
      TabIndex        =   26
      Top             =   3600
      Width           =   2415
   End
   Begin XtremeSuiteControls.Label lblAutorizacion 
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   3165
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Requiere autorización?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aporte Autorizado..:"
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
      Index           =   5
      Left            =   2400
      TabIndex        =   19
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rubro de Patrimonio..:"
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
      Left            =   2400
      TabIndex        =   12
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monto del aporte ..:"
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
      Index           =   2
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmAH_RegistraAhorro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCharRelleno As String, mAporteObreroCal As Currency

Private Sub btnAutorizacion_Click()

On Error GoTo vError

'spPAT_Gestion_Registro(@Cedula varchar(20), @Tipo varchar(10),  @MntCal dec(15,2), @MntSol dec(15,2),  @Usuario varchar(30))
strSQL = "exec spPAT_Gestion_Registro '" & txtCedula.Text & "', '" & cboTipo.ItemData(cboTipo.ListIndex) _
       & "', " & CCur(txtAporteAutorizado.Text) & ", " & CCur(txtMonto.Text) & " , '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)


txtGestionId.Text = rs!Gestion_Id
txtGestionEstado.Text = rs!Gestion_Estado

vError:

End Sub

Private Sub btnRefresh_Click()

On Error GoTo vError

If Not IsNumeric(txtGestionId.Text) Then Exit Sub

strSQL = "exec spPAT_Gestion_Estado " & txtGestionId.Text
Call OpenRecordSet(rs, strSQL)


txtGestionId.Text = rs!Gestion_Id
txtGestionEstado.Text = rs!Gestion_Estado

vError:

End Sub

Private Sub cboTipo_Click()
On Error GoTo vError

Select Case cboTipo.ItemData(cboTipo.ListIndex)
 Case "O" 'Obrero
    txtAporteAutorizado.Text = Format(mAporteObreroCal, "Standard")
 Case "P" 'Patronal
    txtAporteAutorizado.Text = Format(0, "Standard")
 Case "C" 'Capitalizacion
    txtAporteAutorizado.Text = Format(0, "Standard")
 Case "X" 'Custodia
    txtAporteAutorizado.Text = Format(0, "Standard")

End Select

txtGestionId.Text = ""
txtGestionEstado.Text = ""

vError:

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtMonto.SetFocus
End Sub




Private Sub sbAplicar()
Dim vProceso As Long, vCombo As String, curMonto As Currency, i As Integer

Dim vConcepto As String, vTransac As Boolean
Dim vTipoDoc As String, vNumDoc As String

On Error GoTo vError

vTransac = False

Call sbSIFCleanTxtInject(txtNotas)

If Trim(txtNombre.Text) = "" Or Trim(txtMonto.Text) = "" Or Trim(cboTipoDoc.Text) = "" Then
 MsgBox "Faltan Datos", vbExclamation, "No se puede aplicar"
 Exit Sub
End If

If CCur(txtMonto.Text) = 0 Or CCur(txtTotalCajas.Text) = 0 Then
  MsgBox "No se especificó ningún aporte, verifique...", vbExclamation
  Exit Sub
End If


If fxCajasAperturaEstado = "C" Then
   MsgBox "- La apertura ..:" & ModuloCajas.mApertura & " de esta caja ha sido cerrada!", vbExclamation
   Exit Sub
End If

'Verificacion General
If Not fxVerificaDatos Then
        Exit Sub
End If
 
 
 'El aporte es igual a la recaudacion
 txtMonto.Text = txtTotalCajas.Text

 vConcepto = "PAT001"
 
 vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)
 vNumDoc = fxDocumentoConsecutivo(vTipoDoc)
 vProceso = cboProceso.Text

 Call sbDocumento(vTipoDoc, vNumDoc)

 strSQL = "exec spPAT_Aportacion '" & txtCedula.Text & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "'," & CCur(txtMonto.Text) _
        & ",'" & vTipoDoc & "','" & vNumDoc & "','" & glogon.Usuario _
        & "','" & ModuloCajas.mCaja & "','" & vConcepto & "', 0, " & vProceso
 Call ConectionExecute(strSQL)

 Call Bitacora("Registra", "Id: " & txtCedula.Text & ", Pat: " & cboTipo.Text & ", Monto:" & Trim(txtMonto) & ", Doc: " & vTipoDoc & ":" & vNumDoc)
 
 
 If IsNumeric(txtGestionId.Text) And Mid(txtGestionEstado.Text, 1, 1) = "A" And chkAjuste.Value = xtpUnchecked Then
    strSQL = "exec spPAT_Autorizaciones_Aplica " & txtGestionId.Text & ", '" & vTipoDoc & "', '" & vNumDoc & "', '" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
 End If
 
 
 
 'IMPRIMIR RECIBO
If chkReciboDigital.Enabled And chkReciboDigital.Value = xtpChecked Then
    strSQL = "exec spCajasReciboDigital '" & vNumDoc & "', '" & vTipoDoc & "', 'Patrimonio'"
    Call ConectionExecute(strSQL)
    
    MsgBox "Recibo Digital enviado al cliente!", vbInformation

    strSQL = ">>> Recibo Digital enviado al cliente <<<" & vbCrLf _
           & " - Aporte aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
           & " - Desea Realizar Otra Transacción ?"

Else
    Call sbImprimeRecibo(vNumDoc, vTipoDoc)

    strSQL = " - Aporte aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
           & " - Desea Realizar Otra Transacción ?"

End If
Me.MousePointer = vbDefault

 i = MsgBox(strSQL, vbYesNo)
 If i = vbYes Then
     Call sbLimpiaPantalla
 Else
     Unload Me
 End If
 

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCajaInicial()
Dim strSQL As String


'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

pCharRelleno = fxCajasParametros("05")

Me.Caption = "Aportes a Patrimonio       ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

ModuloCajas.mTiquete = "PId." & Trim(txtCedula.Text) & "." & Format(Time, "HH:mm:ss")

If ModuloCajas.mDivisa = "" Then
    ModuloCajas.mDivisa = "COL"
End If

txtTotalCajas.Text = 0
txtNotas.Text = ""

strSQL = "select rTrim(C.tipo_documento) as 'IdX', rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','D')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)


If ModuloCajas.mSesionId = 0 Or ModuloCajas.mClienteId <> ModuloCajas.mSesionCedula Then
   Call sbFormsCall("frmCajas_Sesion", vbModal, , , False, Me)
   If ModuloCajas.mSesionId = 0 Then
        MsgBox "No se ha iniciado ninguna sesión de Cliente para esta caja!", vbExclamation
        Unload Me
        Exit Sub
   End If
End If



End Sub


Private Sub chkAjuste_Click()

If chkAjuste.Value = xtpChecked Then
    lblAutorizacion.Visible = False
    btnRefresh.Visible = False
    btnAutorizacion.Visible = False
Else
    lblAutorizacion.Visible = True
    btnRefresh.Visible = True
    btnAutorizacion.Visible = True
End If

End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtCedula.Text = GLOBALES.gCedulaActual

Dim i As Integer, vProceso As Currency

vProceso = GLOBALES.glngFechaCR

For i = 1 To 6
  cboProceso.AddItem CStr(vProceso)
  vProceso = fxFechaProcesoAnterior(vProceso)
Next i


Call sbLimpiaPantalla

Call Formularios(Me)
Call RefrescaTags(Me)

cboProceso.Text = CStr(GLOBALES.glngFechaCR)

Exit Sub

vError:

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Mantiene o no la Sesion
If ModuloCajas.mSesionId > 0 Then
   Call sbFormsCall("frmCajas_Sesion", vbModal, , , False, Me)
End If

End Sub

Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial


If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Unload Me
   Exit Sub
End If

If ModuloCajas.mSesionId = 0 Or ModuloCajas.mSesionId = Empty Then
   Unload Me
   Exit Sub
End If

End Sub

Private Sub sbLimpiaPantalla()

On Error GoTo vError

strSQL = "select Cedula, Nombre, EstadoActual, (select COD_DIVISA from vSys_Divisas where DIVISA_LOCAL = 1) AS 'COD_DIVISA'" _
       & ", dbo.fxCajas_Valida_Auxiliar('" & ModuloCajas.mCaja & "','PAT','') as 'Caja_Valida_Concepto'" _
       & ", dbo.fxPAT_Info_Aporte_Manual(CEDULA) as 'Pat_Aporte_Manual'" _
       & " from Socios" _
       & " where cedula = '" & txtCedula.Text & "'"
       
Call OpenRecordSet(rs, strSQL)
 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 
 mAporteObreroCal = rs!Pat_Aporte_Manual
 
 ModuloCajas.mCliente = Trim(rs!Nombre)
 ModuloCajas.mClienteId = Trim(rs!Cedula)
 
 ModuloCajas.mReciboDigital = False
 chkReciboDigital.Value = xtpUnchecked
 chkReciboDigital.Enabled = False
 
 txtDivisa.Text = rs!cod_Divisa
 
 ModuloCajas.mDivisa = Trim(rs!cod_Divisa)
 ModuloCajas.mConceptoValida = IIf((rs!Caja_Valida_Concepto > 0), True, False)
 
 
 ModuloCajas.mTiquete = "PId." & Trim(txtCedula.Text) & "." & Format(Time, "HH:mm:ss")
 ModuloCajas.mTotalDetallado = 0
 txtTotalCajas.Text = 0
 
 'FIX
 ModuloCajas.mConceptoValida = True
 
cboTipo.Clear
Select Case rs!EstadoActual
  Case "S"
    cboTipo.AddItem "Aporte Obrero"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "O"
     
    cboTipo.AddItem "Aporte Patronal"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "P"
    
    cboTipo.AddItem "Capitalización"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "C"

    cboTipo.Text = "Aporte Obrero"
 
  Case "A"
    cboTipo.AddItem "Aporte en Custodia"
    cboTipo.ItemData(cboTipo.ListCount - 1) = "X"
    
    cboTipo.Text = "Aporte en Custodia"
  
End Select

rs.Close

txtMonto.Text = "0"
txtNotas.Text = ""

Exit Sub

vError:


End Sub


Private Sub btnCajas_Click(Index As Integer)
On Error GoTo vError

Select Case Index
  Case 2 'Cancelar
     Call sbLimpiaPantalla
     
  Case 0 'Desgloce
        If Not IsNumeric(txtMonto.Text) Then txtMonto.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este Plan/Fondo", vbExclamation
           Exit Sub
        End If
        
        ModuloCajas.mTotalAplicar = CCur(txtMonto.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Patrimonio: Aportaciones"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
        
        If txtTotalCajas.Text <> txtMonto.Text Then
           txtTotalCajas.BackColor = vbRed
        Else
           txtTotalCajas.BackColor = vbWhite
        End If

        If ModuloCajas.mReciboDigital Then
            chkReciboDigital.Enabled = True
            chkReciboDigital.Value = xtpChecked
        Else
            chkReciboDigital.Enabled = False
            chkReciboDigital.Value = xtpUnchecked
        End If

  Case 1  'Aplicar
    Call sbAplicar
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCedula_Change()
txtMonto.Text = 0
txtNombre.Text = ""
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbLimpiaPantalla
End If

End Sub

Private Sub sbAutorizacion_Refresh()
On Error GoTo vError

Select Case cboTipo.ItemData(cboTipo.ListIndex)
 Case "O" 'Obrero
    lblAutorizacion.Tag = "P"
    lblAutorizacion.Caption = "Requiere Autorizacion!"
    txtAporteAutorizado.Text = Format(mAporteObreroCal, "Standard")
 Case "P" 'Patronal
    lblAutorizacion.Tag = "A"
    lblAutorizacion.Caption = "No Requiere Autorizacion!"
    txtAporteAutorizado.Text = Format(0, "Standard")
 Case "C" 'Capitalizacion
    lblAutorizacion.Tag = "A"
    lblAutorizacion.Caption = "No Requiere Autorizacion!"
    txtAporteAutorizado.Text = Format(0, "Standard")
 Case "X" 'Custodia
    lblAutorizacion.Tag = "A"
    lblAutorizacion.Caption = "No Requiere Autorizacion!"
    txtAporteAutorizado.Text = Format(0, "Standard")

End Select


If lblAutorizacion.Tag = "P" Then


End If


vError:

End Sub

Private Function fxVerificaDatos() As Boolean
Dim vMensaje As String

vMensaje = ""

Call sbAutorizacion_Refresh

If Not IsNumeric(txtMonto.Text) Then
  vMensaje = vMensaje & vbCrLf & "- El monto es válido!"
Else
  If CCur(txtMonto.Text) < 0 Then vMensaje = vMensaje & vbCrLf & "- El monto no es válido"
End If

If txtNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & "- El nombre de la persona no es válido"

If cboTipo.Text = "" Then
  vMensaje = vMensaje & vbCrLf & "- No existe tipo de Patrimonio válido para aplicar!"
End If


If Len(vMensaje) = 0 Then
    If chkAjuste.Value = xtpUnchecked Then
      If CCur(txtAporteAutorizado) <> CCur(txtMonto.Text) Then
        If Not IsNumeric(txtGestionId.Text) Or Mid(txtGestionEstado.Text, 1, 1) <> "A" Then
            vMensaje = vMensaje & vbCrLf & "- Este movimiento requiere AUTORIZACION, verifique el estado de la misma y/o solicite una!"
        End If
      End If
    End If
End If

 'Cajas: Validación General sobre el Estado de la Caja, Aperturas, Sesiones, y Accesos
 With ModuloCajas
     strSQL = "exec spCajas_Transac_Validacion '" & .mCaja & "', '" & glogon.Usuario & "', " & .mApertura & ", " & .mSesionId _
            & ", 'PAT', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', " & CCur(txtTotalCajas.Text) & ", '" & .mTiquete & "'"
 End With
 
 Call OpenRecordSet(rs, strSQL)
 
 If Len(rs!Advertencias) > 0 Then
    MsgBox rs!Advertencias, vbExclamation, "Advertencias!"
 End If
 
 If Len(rs!Validacion) > 0 Then
    vMensaje = vMensaje & rs!Validacion
 End If


If Len(vMensaje) = 0 Then
  fxVerificaDatos = True
Else
  fxVerificaDatos = False
  MsgBox vMensaje, vbExclamation
End If

End Function

Private Sub sbDocumento(vTipoDoc As String, vNumDoc As String)

Dim strLinea(10) As String
Dim pDivisa As String, pTipoCambio As Currency
Dim vColCuenta As String, vColAporte As String

On Error GoTo vError

pDivisa = txtDivisa.Text
pTipoCambio = fxCajasTipoCambio(pDivisa)

Select Case cboTipo.ItemData(cboTipo.ListIndex)
  Case "P" 'Aporte Patronal
     vColAporte = "Aporte"
     vColCuenta = "Cta_Patronal"
  
  Case "X" 'Aporte en Custodia"
     vColAporte = "Custodia"
     vColCuenta = "cta_custodia"
 
  Case "O" 'Aporte Obrero"
     vColAporte = "ahorro"
     vColCuenta = "cta_obrero"
  
  Case "C" 'Capitalización"
     vColAporte = "capitaliza"
     vColCuenta = "cta_capitaliza"
End Select


strSQL = "select C." & vColAporte & " as 'Aporte'" _
       & ", (select Cta." & vColCuenta & " as Cuenta from par_afah Cta Where Cta.Cod_Divisa = C.Cod_Divisa) as 'Cuenta'" _
       & " from ahorro_consolidado C" _
       & " where C.cedula = '" & Trim(txtCedula.Text) & "'"

Call OpenRecordSet(rs, strSQL)


strLinea(1) = "Plan            : " & cboTipo.Text
strLinea(2) = "                  "
strLinea(3) = "Saldo Anterior  : " & SIFGlobal.fxStringRelleno(Format(rs!Aporte, "Standard"), "I", pCharRelleno, 20)
strLinea(4) = "Monto del Aporte: " & SIFGlobal.fxStringRelleno(txtMonto.Text, "I", pCharRelleno, 20)
strLinea(5) = "Saldo Actual    : " & SIFGlobal.fxStringRelleno(Format(rs!Aporte + CCur(txtMonto), "Standard"), "I", pCharRelleno, 20)
strLinea(6) = "                  "
strLinea(7) = "Divisa          : " & pDivisa
strLinea(8) = "Caja ¦ Apertura : " & ModuloCajas.mCaja & ": " & ModuloCajas.mApertura
strLinea(9) = "Usuario         : " & glogon.Usuario
strLinea(10) = ""

    
   strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
             & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
             & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento,cod_caja,cod_apertura, id_sesion)" _
             & " values('" & vNumDoc & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" _
             & txtCedula.Text & "','" & txtNombre & "','PAT001'," & CCur(txtMonto) & ",'P','" _
             & txtCedula.Text & "','','','" & GLOBALES.gOficinaTitular & "', " _
             & "'" & strLinea(1) & "','" & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
             & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
             & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
             & vAseDocDetalle & "','" & vAseDocDeposito & "', '" & ModuloCajas.mCaja & "', " & ModuloCajas.mApertura & ", " & ModuloCajas.mSesionId & ")"
    
    'ASIENTO
    If CCur(txtMonto) > 0 Then
        strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & CCur(txtMonto) * fxSys_Tipo_Cambio_Apl(pTipoCambio) & "" _
                & ",'C','" & pDivisa & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & ModuloCajas.mUnidad & "'," _
                & " '','" & rs!Cuenta & "','Pat:" & cboTipo.ItemData(cboTipo.ListIndex) & "','" & txtCedula.Text & "','" & vAseDocDeposito & "'"

        'Procesa Formas de Pago (Registro Final / Asiento de Pago)
        strSQL = strSQL & Space(10) & "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
                & "','" & ModuloCajas.mUsuario & "','" & vTipoDoc & "','" & vNumDoc & "','" & ModuloCajas.mUnidad _
                & "','Pat:" & cboTipo.ItemData(cboTipo.ListIndex) & "','" & txtCedula.Text & "'"
    End If
   
    'Registrar
    Call ConectionExecute(strSQL)

rs.Close

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

Exit Sub

vError:
 
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then btnCajas(0).SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")

Exit Sub

vError:

End Sub
