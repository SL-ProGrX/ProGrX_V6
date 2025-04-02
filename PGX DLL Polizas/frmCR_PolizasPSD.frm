VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCR_PolizasPSD 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Informe: Póliza Saldo Deudores"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6972
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   11052
      _Version        =   1441793
      _ExtentX        =   19494
      _ExtentY        =   12298
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
      SelectedItem    =   1
      Item(0).Caption =   "Informe"
      Item(0).Tooltip =   "Informe al Corte"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "Label2(0)"
      Item(0).Control(2)=   "Label2(1)"
      Item(0).Control(3)=   "dtpCorte"
      Item(0).Control(4)=   "dtpUltimo_Factura"
      Item(0).Control(5)=   "Label2(2)"
      Item(0).Control(6)=   "Label2(3)"
      Item(0).Control(7)=   "btnCorte"
      Item(0).Control(8)=   "Label2(4)"
      Item(0).Control(9)=   "DateTimePicker1"
      Item(0).Control(10)=   "btnExcel(0)"
      Item(0).Control(11)=   "txtFactura_Ultima"
      Item(1).Caption =   "Consulta"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "Opt_Consulta(0)"
      Item(1).Control(1)=   "cboConsulta"
      Item(1).Control(2)=   "vGrid_Corte"
      Item(1).Control(3)=   "btnConsulta(0)"
      Item(1).Control(4)=   "Opt_Consulta(1)"
      Item(1).Control(5)=   "Opt_Consulta(2)"
      Item(1).Control(6)=   "Opt_Consulta(3)"
      Item(1).Control(7)=   "Label2(8)"
      Item(1).Control(8)=   "btnExcel(1)"
      Item(1).Control(9)=   "Opt_Consulta(4)"
      Begin VB.ComboBox cboConsulta 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   1692
      End
      Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
         Height          =   315
         Left            =   -66280
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnCorte 
         Height          =   492
         Left            =   -61840
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Genera Corte"
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
         Picture         =   "frmCR_PolizasPSD.frx":0000
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   315
         Left            =   -68560
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpUltimo_Factura 
         Height          =   315
         Left            =   -68560
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   492
         Index           =   0
         Left            =   -60160
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Excel"
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
         Picture         =   "frmCR_PolizasPSD.frx":0719
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5412
         Left            =   -69880
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   10812
         _Version        =   524288
         _ExtentX        =   19071
         _ExtentY        =   9546
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
         MaxRows         =   1000000
         SpreadDesigner  =   "frmCR_PolizasPSD.frx":0FEA
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   288
         Index           =   0
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Todo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   492
         Index           =   0
         Left            =   8040
         TabIndex        =   15
         Top             =   480
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Carga Información"
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
         Picture         =   "frmCR_PolizasPSD.frx":1723
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   16
         Top             =   480
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Inclusiones"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   2
         Left            =   4200
         TabIndex        =   17
         Top             =   720
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Exclusiones"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   3
         Left            =   5760
         TabIndex        =   18
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Modificaciones"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   492
         Index           =   1
         Left            =   9840
         TabIndex        =   19
         Top             =   480
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Excel"
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
         Picture         =   "frmCR_PolizasPSD.frx":1E2B
      End
      Begin FPSpreadADO.fpSpread vGrid_Corte 
         Height          =   5412
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   10812
         _Version        =   524288
         _ExtentX        =   19071
         _ExtentY        =   9546
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
         MaxRows         =   1000000
         SpreadDesigner  =   "frmCR_PolizasPSD.frx":26FC
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   4
         Left            =   5760
         TabIndex        =   22
         Top             =   720
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Sin Cambios"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtFactura_Ultima 
         Height          =   315
         Left            =   -66280
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
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
         Index           =   8
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
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
         Left            =   -69520
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura:"
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
         Index           =   1
         Left            =   -69520
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura:"
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
         Index           =   2
         Left            =   -67120
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(Referencia para Comparación)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -64240
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Anterior:"
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
         Index           =   4
         Left            =   -67120
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe: Póliza Saldo Deudores"
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
      Height          =   372
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   8772
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11532
   End
End
Attribute VB_Name = "frmCR_PolizasPSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnConsulta_Click(Index As Integer)

Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String


On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case Opt_Consulta.Item(0).Value
     vTipo = "T"
  Case Opt_Consulta.Item(1).Value
     vTipo = "I"
  Case Opt_Consulta.Item(2).Value
     vTipo = "E"
  Case Opt_Consulta.Item(3).Value
     vTipo = "M"
  Case Opt_Consulta.Item(4).Value
     vTipo = "SC"
End Select


strSQL = "exec spPoliza_PSD '', '" & Format(dtpCorte.Value, "yyyy/MM/dd") & "','" & glogon.Usuario & "','" & vTipo & "'"
Call sbCargaGrid(vGrid_Corte, 8, strSQL, True)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbCorte_Consulta(pGrid As Object, pCorte As String, Optional pTipoMov As String = "T")
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbDefault

strSQL = "exec spPolizas_Sicama_Consulta '" & pCorte & "','" & pTipoMov & "'"
Call sbCargaGrid(pGrid, pGrid.MaxCols, strSQL, True)

Exit Sub

vError:
  Me.MousePointer = vbHourglass
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCorte_Genera()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbDefault

strSQL = "exec spPolizas_Sicama_Genera '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbCorte_Consulta(vGrid, Format(dtpCorte.Value, "yyyy/mm/dd"), "T")

Exit Sub

vError:
  Me.MousePointer = vbHourglass
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCorte_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spPoliza_PSD '', '" & Format(dtpCorte.Value, "yyyy/MM/dd") & "','" & glogon.Usuario & "','T'"
Call sbCargaGrid(vGrid, 8, strSQL, True)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExcel_Click(Index As Integer)
Dim vHeaders As vGridHeaders, vFecha As Date, vTipo As String

'Default para Cortes
vHeaders.Columnas = 8
vHeaders.Headers(1) = "Corte"
vHeaders.Headers(2) = "Identificación"
vHeaders.Headers(3) = "Nombre"
vHeaders.Headers(4) = "Monto Asegurado"
vHeaders.Headers(5) = "Fecha Nacimiento"
vHeaders.Headers(6) = "Genero"
vHeaders.Headers(7) = "Nacionalidad"
vHeaders.Headers(8) = "Movimiento"
    

Select Case Index
  Case 0 'Consulta
    Call sbSIFGridExportar(vGrid, vHeaders, "PSD_Corte_" & Format(dtpCorte.Value, "yyyy-mm-dd"))
    
  
  Case 1 'Informe
     Select Case True
       Case Opt_Consulta.Item(0).Value
          vTipo = "TODO"
       Case Opt_Consulta.Item(1).Value
          vTipo = "INCLUSIONES"
       Case Opt_Consulta.Item(2).Value
          vTipo = "EXCLUSIONES"
       Case Opt_Consulta.Item(3).Value
          vTipo = "MODIFICACIONES"
       Case Opt_Consulta.Item(3).Value
          vTipo = "SIN CAMBIOS"
     End Select
     
    Call sbSIFGridExportar(vGrid_Corte, vHeaders, "PSD_Corte_" & cboConsulta.Text & "_" & vTipo)
    
   
End Select

End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


dtpCorte.Value = fxFechaServidor

tcMain.Item(0).Selected = True

vGrid.MaxRows = 0

vGrid_Corte.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next


imgBanner.Width = Me.Width
tcMain.Width = Me.Width - 350
tcMain.Height = Me.Height - (tcMain.Top + 680)
vGrid.Width = tcMain.Width - 250
vGrid_Corte.Width = vGrid.Width

vGrid.Height = tcMain.Height - (vGrid.Top + 250)
vGrid_Corte.Height = tcMain.Height - (vGrid_Corte.Top + 250)

End Sub

Private Sub Opt_Consulta_Click(Index As Integer)

vGrid_Corte.MaxRows = 0

End Sub


