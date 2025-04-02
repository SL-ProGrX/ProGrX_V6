VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmFNDPlanillaFondos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondos: Carga Directa"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11025
   Icon            =   "frmFNDPlanillaFondos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   11025
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   10812
      _Version        =   524288
      _ExtentX        =   19071
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
      SpreadDesigner  =   "frmFNDPlanillaFondos.frx":0ECA
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   2520
      TabIndex        =   8
      Top             =   240
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   2520
      TabIndex        =   9
      Top             =   600
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2520
      TabIndex        =   10
      Top             =   2160
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3413
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   312
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3413
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   312
      Left            =   2520
      TabIndex        =   13
      Top             =   1320
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   432
      Left            =   2520
      TabIndex        =   14
      Top             =   1680
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12086
      _ExtentY        =   762
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   960
      TabIndex        =   15
      Top             =   7200
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
      Top             =   7200
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
   Begin XtremeSuiteControls.FlatEdit txtSocios 
      Height          =   312
      Left            =   3480
      TabIndex        =   17
      Top             =   7200
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
   Begin XtremeSuiteControls.FlatEdit txtContratos 
      Height          =   312
      Left            =   4440
      TabIndex        =   18
      Top             =   7200
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
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   312
      Left            =   2520
      TabIndex        =   23
      Top             =   2880
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
      Height          =   312
      Left            =   4320
      TabIndex        =   24
      Top             =   2880
      Width           =   5052
      _Version        =   1441793
      _ExtentX        =   8911
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
   Begin XtremeSuiteControls.FlatEdit txtComprobante 
      Height          =   312
      Left            =   6600
      TabIndex        =   25
      Top             =   2520
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
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
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   0
      Left            =   9480
      TabIndex        =   26
      Top             =   1680
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmFNDPlanillaFondos.frx":2C3F
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   1
      Left            =   9960
      TabIndex        =   27
      Top             =   1680
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmFNDPlanillaFondos.frx":333F
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   2
      Left            =   10440
      TabIndex        =   28
      Top             =   1680
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmFNDPlanillaFondos.frx":3A58
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   492
      Left            =   6720
      TabIndex        =   29
      Top             =   6960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      Picture         =   "frmFNDPlanillaFondos.frx":4171
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   492
      Left            =   8760
      TabIndex        =   30
      Top             =   6960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
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
      Picture         =   "frmFNDPlanillaFondos.frx":4898
   End
   Begin XtremeSuiteControls.PushButton btnBitacora 
      Height          =   492
      Left            =   8040
      TabIndex        =   31
      Top             =   6960
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   868
      _StockProps     =   79
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
      Picture         =   "frmFNDPlanillaFondos.frx":4F98
   End
   Begin XtremeSuiteControls.CheckBox chkExcel 
      Height          =   615
      Left            =   9600
      TabIndex        =   32
      Top             =   2520
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Archivo Excel"
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
      Appearance      =   16
      Value           =   1
      Alignment       =   1
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
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   6960
      Width           =   972
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
      Top             =   6960
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
      Top             =   7200
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimiento"
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
      Index           =   7
      Left            =   1200
      TabIndex        =   12
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Comprobante"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
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
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
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
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   1335
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
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11292
   End
End
Attribute VB_Name = "frmFNDPlanillaFondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFecha As Date, vPaso As Boolean, mContrato As Long


Private Sub sbLimpia()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vGrid.MaxRows = 0
txtMonto.Text = 0
txtCasos.Text = 0
txtSocios.Text = 0
txtContratos.Text = 0
txtArchivo.Text = ""

  
strSQL = "select dbo.fxFnd_Planillas_Comprobante(" & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ", " & cboProceso.Text & ") as 'NumDoc'"
Call OpenRecordSet(rs, strSQL)
  txtComprobante.Text = rs!NumDoc
rs.Close
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String


strSQL = "select planilla from instituciones" _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
       
Call OpenRecordSet(rs, strSQL)
        
Select Case Index
  
  Case 0 'buscar
  
    txtArchivo.Text = ""
    If chkExcel.Value = vbChecked Then
       Call sbBuscaArchivo(1)
    Else
        Select Case Trim(rs!planilla)
            Case "00", "03"
                Call sbBuscaArchivo(1)
            Case Else
                Call sbBuscaArchivo(2)
        End Select
    End If
  
  Case 1 'Cargar
    If chkExcel.Value = vbChecked Then
       Call sbCargaDeducciones(1)
    Else
         Select Case Trim(rs!planilla)
            Case "00", "03"
               Call sbCargaDeducciones(1)
            Case Else
               Call sbCargaDeducciones(2)
        End Select
    End If
    
  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: FONDOS" & vbCrLf _
              & " 3. Columnas.: CEDULA, NOMBRE, FONDOS"
     
     MsgBox vMensaje, vbInformation
         
End Select

rs.Close

End Sub

Private Sub btnBitacora_Click()
 frmFNDPlanillaBitacora.Show
End Sub

Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub

Private Sub cboInstitucion_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Or cboInstitucion.ListCount = 0 Then Exit Sub
 
Call sbLimpia

If Mid(cboTipo.Text, 1, 1) = "A" Then
    strSQL = "Select cta_Fondos as 'CUENTA' from instituciones where cod_institucion  = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If
 
 
 
If Mid(cboTipo.Text, 1, 1) = "R" Then
    strSQL = "select CUENTA_GASTO as 'CUENTA' From FND_PLANES" _
           & " where COD_OPERADORA = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and COD_PLAN = '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
End If
 
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    txtCuenta.Text = fxgCntCuentaFormato(True, rs!Cuenta, 0)
    txtCuentaDesc.Text = fxgCntCuentaDesc(rs!Cuenta)
Else
    txtCuenta.Text = ""
    txtCuentaDesc.Text = ""
End If
rs.Close
 
 
End Sub

Private Sub cboOperadora_Click()
Dim strSQL As String

If vPaso Or cboOperadora.ListCount = 0 Then Exit Sub

strSQL = "select rtrim(cod_plan) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from fnd_planes where deduce_independiente = 1 and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
vPaso = True

Call sbCbo_Llena_New(cboPlan, strSQL, False, True)

vPaso = False

'Call cboPlan_Click

End Sub



Private Sub cboProceso_Click()
 Call sbLimpia
End Sub

Private Sub cboPlan_Click()
  Call cboInstitucion_Click
End Sub

Private Sub cboTipo_Click()
Call cboInstitucion_Click
End Sub

Private Sub chkExcel_Click()
 Call sbLimpia
End Sub

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer
Dim vProceso As Currency

vModulo = 18

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mFecha = fxFechaServidor

vPaso = True
    strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

    strSQL = "select cod_operadora as IdX, descripcion as ItmX from FND_Operadoras"
    Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

cboTipo.AddItem "Aportes"
cboTipo.AddItem "Rendimientos"
cboTipo.Text = "Aportes"


txtArchivo.Text = ""
txtComprobante.Text = ""

vGrid.MaxCols = 7
vGrid.MaxRows = 0

vProceso = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
vProceso = fxFechaProcesoAnterior(vProceso)


cboProceso.AddItem CStr(vProceso)

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboProceso.AddItem CStr(vProceso)
Next i
cboProceso.Text = CStr(GLOBALES.glngFechaCR)

vPaso = False


Call cboOperadora_Click
Call cboInstitucion_Click


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbCargaDeducciones(vTipo As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

Dim pCedula As String, pNombre As String, pFondos As Currency
Dim pOperadora As Integer, pPlan As String, pInstitucion As Long, pLinea As Long

Dim strCadena As String, curMonto As Currency
Dim fn As Long, lCasos As Long
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

If cboInstitucion.ListCount <= 0 Then
    MsgBox "No existe ninguna Institución, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If
If cboPlan.ListCount <= 0 Then
   MsgBox "No existe ningun plan, no se puede procesar el archivo...", vbCritical
   Exit Sub
End If

If fxAplicada Then
   MsgBox "Ya se aplico una planilla con esta fecha de proceso para la institución y el plan elegidos"
   Exit Sub
End If


Me.MousePointer = vbHourglass


pOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
pPlan = cboPlan.ItemData(cboPlan.ListIndex)
pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)


txtContratos.Text = 0
txtSocios.Text = 0
txtMonto.Text = 0
txtCasos.Text = 0

curMonto = 0
lCasos = 0 'Total

If vTipo = 1 Then 'Archivo de excel

        Set rs = Excel_Load(txtArchivo.Text, "Fondos")
            
        'Validaciónn del Archivo
        vCampos = False
        For i = 0 To rs.Fields.Count
             
            If UCase(LCase(rs.Fields(i).Name)) = "CEDULA" Then
               vCampos = True
            End If
             
             If vCampos Then Exit For
        Next i
        
        If Not vCampos Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Cedula, Nombre, Fondos ¦ Nombre de la Hoja = FONDOS"
           Exit Sub
        End If
        
        vCampos = False
        For i = 0 To rs.Fields.Count
             
            If UCase(LCase(rs.Fields(i).Name)) = "NOMBRE" Then
               vCampos = True
            End If
             
             If vCampos Then Exit For
        Next i
        
        If Not vCampos Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Cedula, Nombre, Fondos ¦ Nombre de la Hoja = FONDOS"
           Exit Sub
        End If
        
        
        vCampos = False
        For i = 0 To rs.Fields.Count
             
            If UCase(LCase(rs.Fields(i).Name)) = "FONDOS" Then
               vCampos = True
            End If
             
             If vCampos Then Exit For
        Next i
        
        If Not vCampos Then
           MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
                 "Los campos son Cedula, Nombre, Fondos ¦ Nombre de la Hoja = FONDOS"
           Exit Sub
        End If
        
        'FIN: Validación del Archivo
        
        
        
        'Sube, Revisa y Carga
        With vGrid
            
            pLinea = 0
            strSQL = ""
            
            Do While Not rs.EOF
              If Trim(rs!Cedula) <> "" Then
                pCedula = rs!Cedula
                pNombre = rs!Nombre
                pFondos = rs!fondos
                pLinea = pLinea + 1
                
                If pLinea = 1 Then
                    strSQL = strSQL & Space(10) & "exec spFndPlanillaDirecta_Sube " & pInstitucion & "," & pOperadora & ",'" & pPlan & "','" _
                           & txtComprobante.Text & "'," & cboProceso.Text & ",'" & pCedula & "','" & pNombre & "'," _
                           & pFondos & "," & pLinea & "," & 1
                Else
                    strSQL = strSQL & Space(10) & "exec spFndPlanillaDirecta_Sube " & pInstitucion & "," & pOperadora & ",'" & pPlan & "','" _
                           & txtComprobante.Text & "'," & cboProceso.Text & ",'" & pCedula & "','" & pNombre & "'," _
                           & pFondos & "," & pLinea & "," & 0
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
        strSQL = "exec spFndPlanillaDirecta_Consulta " & pOperadora & ",'" & pPlan & "','" _
                           & txtComprobante.Text & "',1"
        Call OpenRecordSet(rs, strSQL)
        If glogon.error Then
           Exit Sub
        End If

            Do While Not rs.EOF
                    pCedula = rs!Cedula
                    pNombre = rs!Nombre
                    pFondos = rs!fondos
              
              
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .col = 1
                    .Text = rs!Cedula
                    .col = 2
                    .Text = rs!Nombre
                    .col = 3
                    .Value = IIf((rs!Existe_Persona = 1), 0, 1)
                    
                    .col = 4
                    .Value = IIf((rs!Existe_Contrato = 1), 0, 1)
                    .CellTag = rs!COD_CONTRATO
                    
                    .col = 5
                    .Text = Format(rs!fondos, "Standard")
                    
                    If rs!Existe_Persona = 0 Then
                       txtSocios.Text = CInt(txtSocios.Text) + 1
                    End If
                    
                    If rs!Existe_Contrato = 0 Then
                       txtContratos.Text = CInt(txtContratos.Text) + 1
                    End If
                    
                    curMonto = curMonto + rs!fondos
                    txtMonto.Text = Format(curMonto, "Standard")
                    txtCasos.Text = txtCasos.Text + 1
              
              rs.MoveNext
            Loop
            rs.Close
        
        
    End With 'vGrid


Else 'Archivo Texto
    fn = FreeFile
    Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
    Do While Not EOF(fn)
        Input #fn, strCadena
            
            strNombre = ""
            strCedula = ""
            'monto del archivo
            strMonto = Format(Mid(strCadena, 28, 13), "###########")
            strMonto = LTrim(RTrim(strMonto))
            If Len(strMonto) > 2 Then
                strMonto = Mid(strMonto, 1, Len(strMonto) - 2) & "." & Mid(strMonto, Len(strMonto) - 1, Len(strMonto))
            Else
                strMonto = "0" & "." & strMonto
            End If
            
            curMonto = curMonto + strMonto
            If Len(strCadena) > "54" Then
               strCedula = Trim(Format(Mid(strCadena, 1, 11), "###########"))
               strNombre = Trim(Mid(strCadena, Len(strCadena) - 31, 30))
            End If
            With vGrid
                If Len(strCadena) > "54" Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .col = 1
                    .Text = strCedula
                    .col = 2
                    .Text = strNombre
                    .col = 3
                    If fxNombre(strCedula) = "" Then
                        .Value = 1
                        txtSocios.Text = txtSocios + 1
                    Else
                        .Value = 0
                    End If
                    .col = 4
                    If fxExisteContrato(strCedula) Then
                        .Value = 0
                        txtContratos = txtContratos + 1
                    Else
                        .Value = 1
                    End If
                    .col = 5
                    .Text = Format(strMonto, "Standard")
               End If
            End With
            txtCasos = txtCasos + 1
    Loop
    Close #fn
        
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
Dim strSQL As String
Dim vTipoDoc As String, vNumDoc As String
Dim vCuenta  As String, vInstitucion As Long, vOperadora As Long, vPlan As String

On Error GoTo vError

Me.MousePointer = vbHourglass


vTipoDoc = "PLA"
vNumDoc = txtComprobante.Text
 
vPlan = cboPlan.ItemData(cboPlan.ListIndex)
vInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)
 
strSQL = "exec spFndPlanillaDirecta_Procesa " & vInstitucion & "," & vOperadora & ",'" & vPlan & "'," & cboProceso.Text _
       & ",'" & vNumDoc & "','" & glogon.Usuario & "','" & vCuenta & "','" & Mid(cboTipo.Text, 1, 1) & "'"
Call ConectionExecute(strSQL)
If glogon.error Then
   Exit Sub
End If
 
 
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





Private Function fxAplicada() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select isnull(count(id_seq),0) as Cantidad from fnd_prm_bitacora" _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and cod_plan = '" & SIFGlobal.fxCodText(cboPlan.Text) & "' and proceso = " & cboProceso.Text _
       & " and Documento = '" & txtComprobante.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Cantidad > 0 Then
   fxAplicada = True
Else
   fxAplicada = False
End If
rs.Close

End Function

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
