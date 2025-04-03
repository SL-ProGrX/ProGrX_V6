VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPreaTipos_GastosHonorarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gastos, Honorarios y Examenes: Hipotecarios"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   16365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   16335
      _Version        =   1572864
      _ExtentX        =   28813
      _ExtentY        =   12091
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
      ItemCount       =   4
      Item(0).Caption =   "Bienes Inmuebles"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "cboTipo"
      Item(0).Control(1)=   "vGrid(0)"
      Item(0).Control(2)=   "Label2"
      Item(1).Caption =   "Traspaso"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid(1)"
      Item(2).Caption =   "Exámenes"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGrid(2)"
      Item(3).Caption =   "Avalúo CFIA"
      Item(3).ControlCount=   15
      Item(3).Control(0)=   "Label3(0)"
      Item(3).Control(1)=   "Label3(1)"
      Item(3).Control(2)=   "Label3(2)"
      Item(3).Control(3)=   "Label3(3)"
      Item(3).Control(4)=   "txtCFIA_Formula"
      Item(3).Control(5)=   "txtInterna_Formula"
      Item(3).Control(6)=   "txtIVA_Porc"
      Item(3).Control(7)=   "txtCFIA_HonorarioMinimo"
      Item(3).Control(8)=   "txtR_Usuario"
      Item(3).Control(9)=   "txtR_Fecha"
      Item(3).Control(10)=   "txtA_Usuario"
      Item(3).Control(11)=   "txtA_Fecha"
      Item(3).Control(12)=   "Label3(4)"
      Item(3).Control(13)=   "Label3(5)"
      Item(3).Control(14)=   "btnCFIA_Guardar"
      Begin XtremeSuiteControls.PushButton btnCFIA_Guardar 
         Height          =   495
         Left            =   -65680
         TabIndex        =   21
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Appearance      =   21
         Picture         =   "frmPreaTipos_GastosHonorarios.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   4575
         _Version        =   1572864
         _ExtentX        =   8070
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6135
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   840
         Width           =   16335
         _Version        =   524288
         _ExtentX        =   28813
         _ExtentY        =   10821
         _StockProps     =   64
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
         MaxCols         =   10
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPreaTipos_GastosHonorarios.frx":0731
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6135
         Index           =   1
         Left            =   -70000
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   16335
         _Version        =   524288
         _ExtentX        =   28813
         _ExtentY        =   10821
         _StockProps     =   64
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
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPreaTipos_GastosHonorarios.frx":111E
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6135
         Index           =   2
         Left            =   -70000
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   16335
         _Version        =   524288
         _ExtentX        =   28813
         _ExtentY        =   10821
         _StockProps     =   64
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
         MaxCols         =   12
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPreaTipos_GastosHonorarios.frx":1B81
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCFIA_Formula 
         Height          =   330
         Left            =   -65680
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtInterna_Formula 
         Height          =   330
         Left            =   -65680
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtIVA_Porc 
         Height          =   330
         Left            =   -65680
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCFIA_HonorarioMinimo 
         Height          =   330
         Left            =   -65680
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtA_Usuario 
         Height          =   330
         Left            =   -60280
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtR_Usuario 
         Height          =   330
         Left            =   -60280
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtR_Fecha 
         Height          =   330
         Left            =   -58120
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtA_Fecha 
         Height          =   330
         Left            =   -58120
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   5
         Left            =   -63160
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Actualizado Por"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   4
         Left            =   -63160
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registrado Por"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   3
         Left            =   -69040
         TabIndex        =   10
         Top             =   2280
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto Honorarios Mínimos"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   -69040
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Porcentaje de IVA"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   -69040
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Honorarios Fórmula Interna"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   -69040
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Honorarios Fórmula Crédito Hipotecario"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Concepto"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimientos de Gastos, Honorarios y Examenes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   16455
   End
End
Attribute VB_Name = "frmPreaTipos_GastosHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, mTipo As String

Private Sub sbCFIA_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spCrdPreaListaAvaluoCFIA"
Call OpenRecordSet(rs, strSQL)

txtCFIA_Formula.Text = Format(rs!VALOR_FORMULA_CRD_HIP, "Standard")
txtCFIA_HonorarioMinimo.Text = Format(rs!MONTO_HONORARIOS_MIN_IVA, "Standard")
txtInterna_Formula.Text = Format(rs!VALOR_FORMULA_ASECCSS, "Standard")
txtIVA_Porc = Format(rs!VALOR_PORC_IVA, "Standard")

txtR_Usuario.Text = rs!USUARIO_REGISTRO & ""
txtR_Fecha.Text = rs!FEC_REGISTRO & ""
txtA_Usuario.Text = rs!USUARIO_MODIFICA & ""
txtA_Fecha.Text = rs!FECHA_MODIFICA & ""

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbLista(pTipo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

mTipo = UCase(pTipo)

strSQL = "exec spCrd_Prea_Config_Listas '" & pTipo & "'"

Select Case pTipo
    Case "CANC"
        Call sbCargaGrid(vGrid(0), vGrid(0).MaxCols, strSQL)
    Case "CONS"
        Call sbCargaGrid(vGrid(0), vGrid(0).MaxCols, strSQL)
    Case "TRAS"
        Call sbCargaGrid(vGrid(1), vGrid(1).MaxCols, strSQL)
    Case "EXAM"
        Call sbCargaGrid(vGrid(2), vGrid(2).MaxCols, strSQL)
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCFIA_Guardar_Click()

On Error GoTo vError

strSQL = "exec spCrdPreaModificaAvaluoCFIA " & CCur(txtCFIA_Formula.Text) & ", " & CCur(txtInterna_Formula.Text) _
       & ", " & CCur(txtIVA_Porc.Text) & ", " & CCur(txtCFIA_HonorarioMinimo.Text) & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Información Actualizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub

Call sbLista(UCase(Mid(cboTipo.Text, 1, 4)))

End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
    cboTipo.Clear
    cboTipo.AddItem "Constitución de Hipoteca"
    cboTipo.AddItem "Cancelación de Hipoteca"
    cboTipo.Text = "Constitución de Hipoteca"
vPaso = False


tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

Call cboTipo_Click

End Sub

Private Function fxGuardar(Index As Integer) As Long

On Error GoTo vError

fxGuardar = 0

vGrid(Index).Row = vGrid(Index).ActiveRow
vGrid(Index).Col = 1


Dim pId As Long, pMontoMin As Currency, pMontoMax As Currency, pGastos As Currency, pHonorarios As Currency, pImpuesto As Currency
Dim pRangoEdad As String, pEdadMin As Integer, pEdadMax As Integer, pEdadDesc As String, pEstado As String

pId = 0

With vGrid(Index)

Select Case Index
    Case 0
      .Col = 1
      pId = IIf(.Text = "", 0, .Text)
      .Col = 2
      pMontoMin = CCur(.Text)
      .Col = 3
      pMontoMax = CCur(.Text)
      .Col = 4
      pGastos = CCur(.Text)
      .Col = 5
      pHonorarios = CCur(.Text)
      .Col = 6
      pEstado = Mid(.Text, 1, 1)
    
    Case 1
      .Col = 1
      pId = IIf(.Text = "", 0, .Text)
      .Col = 2
      pMontoMin = CCur(.Text)
      .Col = 3
      pMontoMax = CCur(.Text)
      .Col = 4
      pGastos = CCur(.Text)
      .Col = 5
      pHonorarios = CCur(.Text)
      .Col = 6
      pImpuesto = CCur(.Text)
      .Col = 7
      pEstado = Mid(.Text, 1, 1)
    
    Case 2
      .Col = 1
      pId = IIf(.Text = "", 0, .Text)
      .Col = 2
      pRangoEdad = .Text
      .Col = 3
      pEdadMin = .Text
      .Col = 4
      pEdadMax = .Text
      .Col = 5
      pMontoMin = CCur(.Text)
      .Col = 6
      pMontoMax = CCur(.Text)
      .Col = 7
      pEdadDesc = .Text
      .Col = 8
      pEstado = Mid(.Text, 1, 1)
    
End Select

Select Case mTipo
    Case "CANC"
        strSQL = "exec spCrd_Prea_Config_Hipoteca_Cancelacion_Add " & pId & ", " & pMontoMin & ", " & pMontoMax _
               & ", " & pGastos & ", " & pHonorarios & ", '" & pEstado & "', '" & glogon.Usuario & "'"
    Case "CONS"
        strSQL = "exec spCrd_Prea_Config_Hipoteca_Constitucion_Add " & pId & ", " & pMontoMin & ", " & pMontoMax _
               & ", " & pGastos & ", " & pHonorarios & ", '" & pEstado & "', '" & glogon.Usuario & "'"
    Case "TRAS"
        strSQL = "exec spCrd_Prea_Config_Traspaso_Bienes_Muebles_Add " & pId & ", " & pMontoMin & ", " & pMontoMax _
               & ", " & pGastos & ", " & pHonorarios & ", " & pImpuesto & ", '" & pEstado & "', '" & glogon.Usuario & "'"
    Case "EXAM"
        strSQL = "exec spCrd_Prea_Config_Examen_Requisito_Add " & pId & ", '" & pRangoEdad _
               & "', " & pEdadMin & ", " & pEdadMax & ", " & pMontoMin & ", " & pMontoMax _
               & ", '" & pEdadDesc & "', '" & pEstado & "', '" & glogon.Usuario & "'"
End Select

Call OpenRecordSet(rs, strSQL)
If rs!Pass = 1 Then
  
  pId = rs!IdLlave
  
  .Col = 1
  .Text = rs!IdLlave
  
  Call Bitacora(rs!Movimiento, rs!Mensaje)
  MsgBox rs!Mensaje & ", " & rs!Movimiento & " satisfactoriamente!", vbInformation
Else
   MsgBox rs!Mensaje, vbExclamation
End If


End With

fxGuardar = pId

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 0
        Call sbLista(UCase(Mid(cboTipo.Text, 1, 4)))
    Case 1
        Call sbLista("TRAS")
    Case 2
        Call sbLista("EXAM")
        
    Case 3
        Call sbCFIA_Load
End Select
End Sub

Private Sub vGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
7
Dim i As Long, MaxCol As Integer, Tabla As String

On Error GoTo vError

Select Case Index
    Case 0
        MaxCol = 6
        Tabla = cboTipo.Text
    Case 1
        MaxCol = 7
        Tabla = "Traspaso de Bienes Mueble"
    Case 2
        MaxCol = 8
        Tabla = "Requisitos de Examenes"
End Select


If vGrid(Index).ActiveCol = MaxCol And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar(Index)
  If i = 0 Then Exit Sub
  vGrid(Index).Row = vGrid(Index).ActiveRow
  If vGrid(Index).MaxRows <= vGrid(Index).ActiveRow Then
    vGrid(Index).MaxRows = vGrid(Index).MaxRows + 1
    vGrid(Index).Row = vGrid(Index).MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid(Index).MaxRows = vGrid(Index).MaxRows + 1
    vGrid(Index).InsertRows vGrid(Index).ActiveRow, 1
    vGrid(Index).Row = vGrid(Index).ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar de " & Tabla, vbYesNo)
     If i = vbYes Then
        vGrid(Index).Row = vGrid(Index).ActiveRow
        vGrid(Index).Col = 1
        
        
        Select Case mTipo
            Case "CANC"
                strSQL = "exec spCrd_Prea_Config_Hipoteca_Cancelacion_Del '" & vGrid(Index).Text & "', '" & glogon.Usuario & "'"
            Case "CONS"
                strSQL = "exec spCrd_Prea_Config_Hipoteca_Constitucion_Del '" & vGrid(Index).Text & "', '" & glogon.Usuario & "'"
            Case "TRAS"
                strSQL = "exec spCrd_Prea_Config_Traspaso_Bienes_Muebles_Del '" & vGrid(Index).Text & "', '" & glogon.Usuario & "'"
            Case "EXAM"
                strSQL = "exec spCrd_Prea_Config_Examen_Requisito_Del '" & vGrid(Index).Text & "', '" & glogon.Usuario & "'"
        End Select
        
        
        Call OpenRecordSet(rs, strSQL)
        
        If rs!Pass = 1 Then
                    
            vGrid(Index).Col = 1
            strSQL = vGrid(Index).Text
    
            vGrid(Index).DeleteRows vGrid(Index).ActiveRow, 1
            vGrid(Index).MaxRows = vGrid(Index).MaxRows - 1
            
            If vGrid(Index).MaxRows <= 0 Then
              vGrid(Index).MaxRows = 1
            End If
            
            Call Bitacora(rs!Movimiento, rs!Mensaje)
            
            MsgBox rs!Mensaje & ", Eliminado Satisfactoriamente!", vbInformation
        Else
            MsgBox rs!Mensaje, vbExclamation
        End If


     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

