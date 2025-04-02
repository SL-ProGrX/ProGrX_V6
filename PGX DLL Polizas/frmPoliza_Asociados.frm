VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmPoliza_Asociados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Póliza de Vida Colectiva: Asociados"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   11055
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
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Informe"
      Item(0).Tooltip =   "Informe al Corte"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "Label2(0)"
      Item(0).Control(2)=   "dtpCorte"
      Item(0).Control(3)=   "btnCorte"
      Item(0).Control(4)=   "btnExcel(0)"
      Item(1).Caption =   "Consulta"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "Opt_Consulta(0)"
      Item(1).Control(1)=   "btnConsulta(0)"
      Item(1).Control(2)=   "Opt_Consulta(1)"
      Item(1).Control(3)=   "Opt_Consulta(2)"
      Item(1).Control(4)=   "Label2(8)"
      Item(1).Control(5)=   "btnExcel(1)"
      Item(1).Control(6)=   "vGrid_Corte"
      Item(1).Control(7)=   "dtpCorte_Consulta"
      Item(1).Control(8)=   "Opt_Consulta(3)"
      Item(2).Caption =   "Beneficiarios"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "btnConsulta(1)"
      Item(2).Control(1)=   "btnExcel(2)"
      Item(2).Control(2)=   "cboPoliza"
      Item(2).Control(3)=   "vGrid_Beneficiarios"
      Begin XtremeSuiteControls.PushButton btnCorte 
         Height          =   495
         Left            =   -61840
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
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
         Picture         =   "frmPoliza_Asociados.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   315
         Left            =   -68560
         TabIndex        =   3
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
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   495
         Index           =   0
         Left            =   -60160
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
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
         Picture         =   "frmPoliza_Asociados.frx":0719
         ImageAlignment  =   4
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5775
         Left            =   -69880
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   10186
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
         MaxCols         =   19
         MaxRows         =   1000000
         SpreadDesigner  =   "frmPoliza_Asociados.frx":0FEA
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   405
         Index           =   0
         Left            =   3000
         TabIndex        =   6
         Top             =   480
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   714
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
         TabIndex        =   7
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
         Picture         =   "frmPoliza_Asociados.frx":1933
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   8
         Top             =   360
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
         TabIndex        =   9
         Top             =   720
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   503
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
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   492
         Index           =   1
         Left            =   9840
         TabIndex        =   10
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
         Picture         =   "frmPoliza_Asociados.frx":203B
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   405
         Index           =   3
         Left            =   5640
         TabIndex        =   11
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   714
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte_Consulta 
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   600
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
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   495
         Index           =   1
         Left            =   -65080
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
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
         Picture         =   "frmPoliza_Asociados.frx":290C
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   495
         Index           =   2
         Left            =   -63280
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
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
         Picture         =   "frmPoliza_Asociados.frx":3014
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboPoliza 
         Height          =   465
         Left            =   -70000
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
         _Version        =   1441793
         _ExtentX        =   8493
         _ExtentY        =   820
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         FlatStyle       =   -1  'True
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread vGrid_Beneficiarios 
         Height          =   5775
         Left            =   -70000
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   10186
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
         MaxCols         =   32
         MaxRows         =   1000000
         SpreadDesigner  =   "frmPoliza_Asociados.frx":38E5
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid_Corte 
         Height          =   5775
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   10186
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
         MaxCols         =   19
         MaxRows         =   1000000
         SpreadDesigner  =   "frmPoliza_Asociados.frx":4A8A
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
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
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   852
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
         TabIndex        =   12
         Top             =   600
         Width           =   852
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe: Póliza de Vida Colectiva: Asociados"
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
Attribute VB_Name = "frmPoliza_Asociados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub btnConsulta_Click(Index As Integer)
Dim vTipo As String


On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Index
  Case 0
        Select Case True
          Case Opt_Consulta.Item(0).Value
             vTipo = "T"
          Case Opt_Consulta.Item(1).Value
             vTipo = "I"
          Case Opt_Consulta.Item(2).Value
             vTipo = "E"
          Case Opt_Consulta.Item(3).Value
             vTipo = "SC"
        End Select
        
        
        strSQL = "exec spPoliza_Asociados '" & Format(dtpCorte_Consulta.Value, "yyyy/MM/dd") & "','" & glogon.Usuario & "','" & vTipo & "'"
        Call sbCargaGrid(vGrid_Corte, 19, strSQL, True)
  
  Case 1 'Consulta de Beneficiarios
    Call sbConsulta_Beneficiarios

End Select

Me.MousePointer = vbDefault

Exit Sub

vError:


End Sub

Private Sub sbConsulta_Beneficiarios()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbDefault

strSQL = "exec spPoliza_Beneficiarios_Lista '" & cboPoliza.ItemData(cboPoliza.ListIndex) & "'"
Call sbCargaGrid(vGrid_Beneficiarios, vGrid_Beneficiarios.MaxCols, strSQL, True)

Exit Sub

vError:
  Me.MousePointer = vbHourglass
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub btnCorte_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPoliza_Asociados '" & Format(dtpCorte.Value, "yyyy/MM/dd") & "','" & glogon.Usuario & "','T'"
Call sbCargaGrid(vGrid, 19, strSQL, True)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExcel_Click(Index As Integer)
Dim vHeaders As vGridHeaders, vFecha As Date, vTipo As String

'Default para Cortes
vHeaders.Columnas = 19
vHeaders.Headers(1) = "Corte"
vHeaders.Headers(2) = "Identificación"
vHeaders.Headers(3) = "Id Alterno"
vHeaders.Headers(4) = "Apellido_1"
vHeaders.Headers(5) = "Apellido_2"
vHeaders.Headers(6) = "Nombre_1"
vHeaders.Headers(7) = "Nombre_2"

vHeaders.Headers(8) = "Email_1"
vHeaders.Headers(9) = "Email_2"
vHeaders.Headers(10) = "Fecha Nacimiento"
vHeaders.Headers(11) = "Genero"
vHeaders.Headers(12) = "Nacionalidad"

vHeaders.Headers(13) = "Provincia"
vHeaders.Headers(14) = "Cantón"
vHeaders.Headers(15) = "Distrito"
vHeaders.Headers(16) = "Dirección"
vHeaders.Headers(17) = "Tipo Teléfono"
vHeaders.Headers(18) = "Num. Teléfono"


vHeaders.Headers(19) = "Movimiento"
    

Select Case Index
  Case 0 'Consulta
    Call sbSIFGridExportar(vGrid, vHeaders, "Poliza_PA_Asociados_Corte_" & Format(dtpCorte.Value, "yyyy-mm-dd"))
    
  
  Case 1 'Informe
  
  
     Select Case True
       Case Opt_Consulta.Item(0).Value
          vTipo = "TODO"
       Case Opt_Consulta.Item(1).Value
          vTipo = "INCLUSIONES"
       Case Opt_Consulta.Item(2).Value
          vTipo = "EXCLUSIONES"
       Case Opt_Consulta.Item(3).Value
          vTipo = "SIN_CAMBIOS"
       Case Opt_Consulta.Item(4).Value
          vTipo = "SIN_CAMBIOS"
     End Select
     
    Call sbSIFGridExportar(vGrid_Corte, vHeaders, "Poliza_PA_Asociados_Corte_" & Format(dtpCorte_Consulta.Value, "yyyy-mm-dd") & "_" & vTipo)
    
   
   
  Case 2 'Beneficiarios
    vFecha = fxFechaServidor
    
    vHeaders.Columnas = vGrid_Beneficiarios.MaxCols
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Bene No.1 Tipo ( ID )"
    vHeaders.Headers(4) = "Bene No.1 Identificación"
    vHeaders.Headers(5) = "Bene No.1 Nombre Completo"
    vHeaders.Headers(6) = "Bene No.1 Parentesco"
    vHeaders.Headers(7) = "Bene No.1 Porcentaje"
    
    vHeaders.Headers(8) = "Bene No.2 Tipo ( ID )"
    vHeaders.Headers(9) = "Bene No.2 Identificación"
    vHeaders.Headers(10) = "Bene No.2 Nombre Completo"
    vHeaders.Headers(11) = "Bene No.2 Parentesco"
    vHeaders.Headers(12) = "Bene No.2 Porcentaje"
    
    vHeaders.Headers(13) = "Bene No.3 Tipo ( ID )"
    vHeaders.Headers(14) = "Bene No.3 Identificación"
    vHeaders.Headers(15) = "Bene No.3 Nombre Completo"
    vHeaders.Headers(16) = "Bene No.3 Parentesco"
    vHeaders.Headers(17) = "Bene No.3 Porcentaje"
    
    
    vHeaders.Headers(18) = "Bene No.4 Tipo ( ID )"
    vHeaders.Headers(19) = "Bene No.4 Identificación"
    vHeaders.Headers(20) = "Bene No.4 Nombre Completo"
    vHeaders.Headers(21) = "Bene No.4 Parentesco"
    vHeaders.Headers(22) = "Bene No.4 Porcentaje"
    
    vHeaders.Headers(23) = "Bene No.5 Tipo ( ID )"
    vHeaders.Headers(24) = "Bene No.5 Identificación"
    vHeaders.Headers(25) = "Bene No.5 Nombre Completo"
    vHeaders.Headers(26) = "Bene No.5 Parentesco"
    vHeaders.Headers(27) = "Bene No.5 Porcentaje"
    
    vHeaders.Headers(28) = "Bene No.6 Tipo ( ID )"
    vHeaders.Headers(29) = "Bene No.6 Identificación"
    vHeaders.Headers(30) = "Bene No.6 Nombre Completo"
    vHeaders.Headers(31) = "Bene No.6 Parentesco"
    vHeaders.Headers(32) = "Bene No.6 Porcentaje"
    
    
    Call sbSIFGridExportar(vGrid_Beneficiarios, vHeaders, "Poliza_PA_Asociados_Beneficiarios_" & Format(vFecha, "yyyy-mm-dd"))
   
End Select

End Sub


Private Sub Form_Load()

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


 strSQL = "select COD_POLIZA as 'Idx', rtrim(Poliza_Desc) as 'ItmX' from vPoliza_Catalogo" _
        & " Where Tipo = 'PA'" _
        & " order by COD_POLIZA"
 Call sbCbo_Llena_New(cboPoliza, strSQL, False, True)


dtpCorte.Value = fxFechaServidor
dtpCorte_Consulta.Value = dtpCorte.Value

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

vGrid_Beneficiarios.Height = tcMain.Height - (vGrid_Beneficiarios.Top + 250)
vGrid_Beneficiarios.Width = vGrid.Width

End Sub

Private Sub Opt_Consulta_Click(Index As Integer)

vGrid_Corte.MaxRows = 0

End Sub

