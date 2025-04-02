VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCajas_FNDAportaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cajas: Fondos"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9000
      Top             =   240
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   1212
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   8412
      _Version        =   524288
      _ExtentX        =   14838
      _ExtentY        =   2138
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmCajas_FNDAportaciones.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1572
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   9492
      _Version        =   1441793
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
         TabIndex        =   11
         Top             =   240
         Width           =   2772
         _Version        =   1441793
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
         Height          =   312
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   5412
         _Version        =   1441793
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
         Height          =   792
         Index           =   0
         Left            =   6720
         TabIndex        =   14
         Top             =   600
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Pago"
         BackColor       =   -2147483633
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
         Picture         =   "frmCajas_FNDAportaciones.frx":05EE
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   792
         Index           =   1
         Left            =   7680
         TabIndex        =   15
         Top             =   600
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Aplicar"
         BackColor       =   -2147483633
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
         Picture         =   "frmCajas_FNDAportaciones.frx":0A9B
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   792
         Index           =   2
         Left            =   8520
         TabIndex        =   16
         Top             =   600
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Cancelar"
         BackColor       =   -2147483633
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
         Picture         =   "frmCajas_FNDAportaciones.frx":1273
         TextImageRelation=   1
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
         TabIndex        =   19
         Top             =   240
         Width           =   1452
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
         TabIndex        =   18
         Top             =   600
         Width           =   1452
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
         TabIndex        =   17
         Top             =   240
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3480
      TabIndex        =   20
      Top             =   240
      Width           =   5292
      _Version        =   1441793
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
      Left            =   1800
      TabIndex        =   21
      Top             =   240
      Width           =   1692
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMensualidad 
      Height          =   312
      Left            =   5160
      TabIndex        =   22
      Top             =   3480
      Width           =   1692
      _Version        =   1441793
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
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAporte 
      Height          =   312
      Left            =   5160
      TabIndex        =   23
      Top             =   3840
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   492
      Left            =   8160
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCajas_FNDAportaciones.frx":1A40
   End
   Begin XtremeSuiteControls.FlatEdit txtContrato 
      Height          =   315
      Left            =   1800
      TabIndex        =   25
      Top             =   960
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   315
      Left            =   7800
      TabIndex        =   26
      Top             =   1680
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   315
      Left            =   7800
      TabIndex        =   27
      Top             =   1320
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   1800
      TabIndex        =   28
      Top             =   1320
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
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
   Begin XtremeSuiteControls.FlatEdit txtOperadora 
      Height          =   315
      Left            =   1800
      TabIndex        =   29
      Top             =   1680
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
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
   Begin XtremeSuiteControls.FlatEdit txtDivisa 
      Height          =   315
      Left            =   6840
      TabIndex        =   30
      Top             =   3840
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensualidad ..:"
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
      Index           =   5
      Left            =   3600
      TabIndex        =   9
      Top             =   3480
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Aporte ..:"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   8
      Top             =   3840
      Width           =   1452
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
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label3 
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
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6840
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6840
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SubCuentas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmCajas_FNDAportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCharRelleno As String

Private Sub btnCajas_Click(Index As Integer)
On Error GoTo vError

Select Case Index
  Case 2 'Cancelar
     Call sbLimpiaPantalla
     
  Case 0 'Desgloce
        If Not IsNumeric(txtAporte.Text) Then txtAporte.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este Plan/Fondo", vbExclamation
           Exit Sub
        End If
        
        ModuloCajas.mTotalAplicar = CCur(txtAporte.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Fondos: Aportaciones"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")

  Case 1  'Aplicar
    Call CmdAplicar_Click
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
vError:
End Sub



Public Sub sbDocumento(vTipoDoc As String, vNumDoc As String, ByVal vOperadora As Long, vPlan As String, vContrato As Long _
                              , pAportes As Currency, pConcepto As String, Optional pRendimiento As Currency = 0)
                              
Dim strSQL As String, rs As New ADODB.Recordset, strLinea(11) As String
Dim curMonto As Currency, i As Integer, vTipo As String
Dim vCuentaFND As String, vConcepto As String
Dim vCuentaRendi As String, strDivisa As String, pTipoCambio As Currency



vConcepto = vConcepto & vPlan & " OP:" & vOperadora & " CNT:" & vContrato

strDivisa = fxFndDivisa(vOperadora, vPlan)
pTipoCambio = fxCajasTipoCambio(strDivisa)

rs.Open "select cuenta_conta,cuenta_rendimiento from fnd_planes where cod_operadora = " & vOperadora _
  & " and cod_plan = '" & vPlan & "'", glogon.Conection, adOpenStatic
    vCuentaFND = Trim(rs!Cuenta_Conta)
    vCuentaRendi = Trim(rs!Cuenta_Rendimiento)
rs.Close


strSQL = "Select F.Aportes,F.Rendimiento, S.Cedula,S.Nombre From Fnd_Contratos F inner join Socios S on F.cedula = S.cedula " _
         & "Where F.cod_operadora = " & vOperadora & " and F.cod_plan='" & vPlan & "' and F.cod_contrato=" & vContrato
         
Call OpenRecordSet(rs, strSQL)
  If rs.EOF = False Then
     curMonto = IIf(IsNull(rs!APORTES), 0, rs!APORTES) + IIf(IsNull(rs!Rendimiento), 0, rs!Rendimiento)
  End If
rs.Close

strLinea(1) = "MNT. ANTERIOR   : " & SIFGlobal.fxStringRelleno(Format(curMonto, "Standard"), "I", pCharRelleno, 20)
strLinea(2) = "APORTE APLICADO : " & SIFGlobal.fxStringRelleno(Format(pAportes, "Standard"), "I", pCharRelleno, 20)
strLinea(3) = "RENDI. APLICADO : " & SIFGlobal.fxStringRelleno(Format(pRendimiento, "Standard"), "I", pCharRelleno, 20)
strLinea(4) = "MNT. ACTUAL     : " & SIFGlobal.fxStringRelleno(Format(curMonto + pAportes + pRendimiento, "Standard"), "I", pCharRelleno, 20)
strLinea(5) = "DIVISA : " & strDivisa
strLinea(6) = "[Plan.:" & vPlan & " Cnt.:" & vContrato & "]"
strLinea(7) = "[Operadora.: " & txtOperadora.Text & "]"

'Control de Documentos v2
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,cod_oficina" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7, detalle,documento,cod_caja,cod_apertura)" _
        & " values('" & vNumDoc & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtCedula.Text _
        & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & CCur(pAportes + pRendimiento) & ",'P','" & vPlan _
        & "','" & gFondos.Contrato & "','" & ModuloCajas.mOficina & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & txtNotas.Text & "','" & vAseDocDeposito & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ")"
'Call ConectionExecute(strSQL)


strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & CCur(pAportes) * fxSys_Tipo_Cambio_Apl(pTipoCambio) & "" _
        & ",'C','" & strDivisa & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & ModuloCajas.mUnidad & "'," _
        & " '','" & vCuentaFND & "','" & vPlan & "','" & gFondos.Contrato & "','" & vAseDocDeposito & "'"
'Call ConectionExecute(strSQL)

If pRendimiento > 0 Then
    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & CCur(pRendimiento) * fxSys_Tipo_Cambio_Apl(pTipoCambio) & "" _
            & ",'C','" & strDivisa & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & ModuloCajas.mUnidad & "'," _
            & " '" & ModuloCajas.mCentroCosto & "','" & vCuentaRendi _
            & "','" & vPlan & "','" & gFondos.Contrato & "','" & vAseDocDeposito & "'"
'    Call ConectionExecute(strSQL)
End If

'Procesa Formas de Pago (Registro Final / Asiento de Pago)
strSQL = strSQL & Space(10) & "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
        & "','" & ModuloCajas.mUsuario & "','" & vTipoDoc & "','" & vNumDoc & "','" & ModuloCajas.mUnidad _
        & "','" & vPlan & "','" & gFondos.Contrato & "'"
'Call ConectionExecute(strSQL)


'Aplica en una sola Llamada
Call ConectionExecute(strSQL)

End Sub

Private Sub CmdAplicar_Click()
Dim vFecha As Date, i As Integer
Dim vProceso As Long, vCombo As String, curMonto As Currency

Dim strSQL As String, vConcepto As String, vTransac As Boolean
Dim vTipoDoc As String, vNumDoc As String

On Error GoTo vError


vTransac = False

Call sbSIFCleanTxtInject(txtNotas)

If txtAporte.Tag = "N" Then
  MsgBox "Este plan no permite movimientos en Cajas, verifique...", vbExclamation
  Exit Sub
End If

If txtEstado.Tag = "L" Then
  MsgBox "Este contrato se encuentra Liquidado, verifique...", vbExclamation
  Exit Sub
End If

If CCur(txtAporte) = 0 Or CCur(txtTotalCajas.Text) = 0 Then
  MsgBox "No se especificó ningún aporte, verifique...", vbExclamation
  Exit Sub
End If

If Trim(txtContrato) = "" Or Trim(txtAporte) = "" Or Trim(cboTipoDoc.Text) = "" Then
 MsgBox "Faltan Datos", vbExclamation, "No se puede aplicar"
 Exit Sub
Else

If fxCajasAperturaEstado = "C" Then
   MsgBox "- La apertura ..:" & ModuloCajas.mApertura & " de esta caja ha sido cerrada!", vbExclamation
   Exit Sub
End If

 
If fxFndParametro("01.1") = "S" Then
   strSQL = "exec spFndSeguridad_ApAnul " & gFondos.Operadora & ",'" & gFondos.Plan & "','" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
   If rs!autoriza = 0 Then
        MsgBox "El Usuario no tiene nivel de Autorización para realizar este movimiento!", vbExclamation
        Exit Sub
   End If
End If
     
 
 
 
 'El aporte es igual a la recaudacion
 txtAporte.Text = txtTotalCajas.Text

 vConcepto = "FND001"
 vFecha = fxFechaServidor
 vProceso = Year(vFecha) & Format(Month(vFecha), "00")
 vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)
 
 vNumDoc = fxDocumentoConsecutivo(vTipoDoc)

 Call sbDocumento(vTipoDoc, vNumDoc, gFondos.Operadora, Trim(gFondos.Plan), Trim(txtContrato), CCur(txtAporte), vConcepto, 0)
  
 strSQL = "Insert fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato,Fecha,Monto,Fecha_Proceso,Tcon,Ncon" _
        & ",cod_concepto,usuario,cod_Caja) Values(" & gFondos.Operadora & ",'" _
        & Trim(gFondos.Plan) & "'," & gFondos.Contrato & ",dbo.MyGetdate()," _
        & CCur(txtAporte) & "," & vProceso & ",'" & vTipoDoc & "','" & vNumDoc & "','" & vConcepto _
        & "','" & glogon.Usuario & "','" & ModuloCajas.mCaja & "')"
'        Call ConectionExecute(strSQL)
 
 strSQL = strSQL & Space(10) & "Update Fnd_contratos set Aportes = Aportes + " & CCur(txtAporte) _
        & " where cod_operadora=" & gFondos.Operadora _
        & " and cod_plan='" & Trim(gFondos.Plan) & "'" _
        & " and cod_contrato=" & gFondos.Contrato
'        Call ConectionExecute(strSQL)
 
 
 If txtAporte.Locked Then
  For i = 1 To vGrid.MaxRows
     vGrid.col = 4
     vGrid.Row = i
     
     curMonto = CCur(vGrid.Text)
     vGrid.col = 1
     If curMonto > 0 Then
     
        strSQL = strSQL & Space(10) & "Insert into fnd_SubCuentas_detalle(idx,Cod_operadora,Cod_plan,Cod_Contrato," _
               & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon) Values(" & vGrid.Text & "," & gFondos.Operadora & ",'" _
               & Trim(gFondos.Plan) & "'," & gFondos.Contrato & ",dbo.MyGetdate()," _
               & curMonto & "," & vProceso & ",'" & vTipoDoc & "','" & vNumDoc & "')"
'        Call ConectionExecute(strSQL)
        
        strSQL = strSQL & Space(10) & "Update Fnd_subCuentas set Aportes = Aportes + " & curMonto _
               & " where cod_operadora=" & gFondos.Operadora _
               & " and cod_plan='" & Trim(gFondos.Plan) & "'" _
               & " and cod_contrato=" & gFondos.Contrato & " and IdX = " & vGrid.Text
'        Call ConectionExecute(strSQL)
     
     End If
  Next i
 End If
 
 
 glogon.Conection.BeginTrans
 vTransac = True
 
 'Aplica en una sola llamada
 Call ConectionExecute(strSQL)
 
 glogon.Conection.CommitTrans
 vTransac = False
 
 Call Bitacora("Registra", vTipoDoc & " Ope:" & gFondos.Operadora & " Plan:" & Trim(gFondos.Plan) & " Cont:" & Trim(txtContrato) & " Monto:" & Trim(txtAporte))
 
 Call sbImprimeRecibo(vNumDoc, vTipoDoc)
 
 strSQL = " - Aporte aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
        & " - Desea Realizar Otra Transacción a Este Contrato ?"
 
 i = MsgBox(strSQL, vbYesNo)
 If i = vbYes Then
     Call sbLimpiaPantalla
 Else
     Unload Me
 End If
 
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  If vTransac Then glogon.Conection.RollbackTrans

End Sub

Private Sub Form_Activate()
vModulo = 5

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

Me.Caption = "Aportes a Fondos       ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

ModuloCajas.mTiquete = Trim(gFondos.Plan) & "." & gFondos.Contrato & "." & Format(Time, "HH:mm:ss")

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
End Sub



Private Sub Form_Load()
 
vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle

Call sbLimpiaPantalla

Call Formularios(Me)
Call RefrescaTags(Me)

btnCajas.Item(1).Enabled = cmdAplicar.Enabled

End Sub


Private Sub sbLimpiaPantalla()
Dim strSQL As String, rs As New ADODB.Recordset

txtAporte.Text = 0
txtAporte.Tag = "S"
txtAporte.Locked = False

ModuloCajas.mTiquete = Trim(gFondos.Plan) & "." & gFondos.Contrato & "." & Format(Time, "HH:mm:ss")
ModuloCajas.mTotalDetallado = 0
txtTotalCajas.Text = 0

vGrid.MaxRows = 0

strSQL = "select C.cedula,S.nombre,P.descripcion as PlanX,O.descripcion as OperadoraX,C.Monto" _
       & ",C.cod_plan,C.cod_contrato,C.cod_operadora,C.estado,C.fecha_Inicio,isnull(P.cuenta_Maestra,0) as CuentaMaestra" _
       & ",P.Tipo_CDP,C.Inversion,P.Permite_Mov_Cajas,C.aportes,P.cod_Moneda" _
       & ",dbo.fxCajas_Valida_Auxiliar('" & ModuloCajas.mCaja & "','FND',C.cod_Plan) as 'Caja_Valida_Concepto'" _
       & " from fnd_contratos C inner join Socios S on C.cedula = S.cedula" _
       & " inner join fnd_planes P on C.cod_plan = P.cod_plan and C.cod_operadora = P.cod_operadora" _
       & " inner join fnd_operadoras O on C.cod_operadora = O.cod_operadora" _
       & " where C.cod_operadora = " & gFondos.Operadora & " and C.cod_plan = '" & gFondos.Plan _
       & "' and C.cod_Contrato = " & gFondos.Contrato
       
Call OpenRecordSet(rs, strSQL)
 txtCedula = rs!Cedula
 txtNombre = rs!Nombre
 
 ModuloCajas.mCliente = Trim(rs!Nombre)
 ModuloCajas.mClienteId = Trim(rs!Cedula)
 
 ModuloCajas.mDivisa = Trim(rs!Cod_Moneda)
 ModuloCajas.mConceptoValida = IIf((rs!Caja_Valida_Concepto > 0), True, False)
 
 ModuloCajas.mTotalDetallado = 0
 txtTotalCajas.Text = 0
 
 txtDivisa.Text = Trim(rs!Cod_Moneda)
 txtOperadora.Text = rs!operadoraX
 txtDescripcion.Text = rs!PlanX
 txtContrato.Text = rs!COD_CONTRATO
 txtFecha.Text = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
 txtEstado.Tag = rs!Estado
 txtEstado.Text = IIf((rs!Estado = "A"), "Activo", "Liquidado")
 txtMensualidad.Text = Format(rs!Monto, "Standard")
 
 If rs!tipo_cdp = 1 Then
    txtAporte.Locked = True
    If rs!APORTES = 0 Then
       txtAporte = Format(rs!Inversion, "Standard")
    End If
 End If
 
 
 If rs!CuentaMaestra = 1 Then
    txtAporte.Locked = True
    strSQL = "select IDx,Cedula,Nombre,0 from fnd_subCuentas where cod_operadora = " & gFondos.Operadora _
           & " and cod_plan = '" & gFondos.Plan & "' and cod_contrato = " & gFondos.Contrato _
           & " and estado = 'A'"
   Call sbCargaGrid(vGrid, 4, strSQL)
   vGrid.MaxRows = vGrid.MaxRows - 1
 End If
 
 If rs!PERMITE_MOV_CAJAS = 0 Then
  txtAporte.Tag = "N"
 End If
rs.Close

End Sub



Private Sub tblDesgloce_ButtonClick(ByVal Button As MSComctlLib.Button)

If Not IsNumeric(txtAporte.Text) Then txtAporte.Text = 0

ModuloCajas.mTotalAplicar = CCur(txtAporte.Text)

If ModuloCajas.mTotalAplicar = 0 Then
    MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
    Exit Sub
End If

ModuloCajas.mServicio = "Fondos: Aportaciones"

Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)

txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
    
End Sub



Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial
End Sub


Private Sub txtAporte_GotFocus()
On Error GoTo vError
 txtAporte = CCur(txtAporte)
vError:
End Sub

Private Sub txtAporte_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8, 46
  Case vbKeyReturn
     cboTipoDoc.SetFocus
  Case Else
     KeyAscii = 0
End Select
End Sub


Private Sub txtAporte_LostFocus()
On Error GoTo vError
 txtAporte = Format(CCur(txtAporte), "Standard")
vError:
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAporte.SetFocus
End Sub


Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim i As Integer, curMonto As Currency

If col = 4 Then
 curMonto = 0
 For i = 1 To vGrid.MaxRows
   vGrid.col = 4
   vGrid.Row = i
   curMonto = curMonto + CCur(vGrid.Text)
 Next i
 
 txtAporte = Format(curMonto, "Standard")
 
End If

End Sub


