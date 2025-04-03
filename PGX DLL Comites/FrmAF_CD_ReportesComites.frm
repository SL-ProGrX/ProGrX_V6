VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAF_CD_ReportesComites 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Comites y Delegados"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   Icon            =   "FrmAF_CD_ReportesComites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   495
      Left            =   8160
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Reporte"
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
      Picture         =   "FrmAF_CD_ReportesComites.frx":3482
      ImageAlignment  =   4
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":3B89
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":3CA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":3DB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":3EDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   4080
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7197
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgArbol"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   4800
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":4004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A866
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A984
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":AAAE
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":ABD4
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":ACE2
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":ADEF
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":AF08
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":B036
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":B143
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   5520
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   6960
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   5520
      TabIndex        =   13
      Top             =   1800
      Width           =   2895
      _Version        =   1572864
      _ExtentX        =   5106
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
   Begin XtremeSuiteControls.ComboBox cboPromotores 
      Height          =   330
      Left            =   5520
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7646
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
   Begin XtremeSuiteControls.ComboBox cboComite 
      Height          =   330
      Left            =   5520
      TabIndex        =   15
      Top             =   3000
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7646
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
   Begin XtremeSuiteControls.ComboBox cboActividades 
      Height          =   330
      Left            =   5520
      TabIndex        =   16
      Top             =   3480
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7646
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   5520
      TabIndex        =   17
      Top             =   3960
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
   Begin XtremeSuiteControls.PushButton btnInformeEspecial 
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   5280
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Informe Especial de Antiguedad"
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
      Picture         =   "FrmAF_CD_ReportesComites.frx":B243
      ImageAlignment  =   4
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   4215
      _Version        =   1572864
      _ExtentX        =   7435
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Informes Disponibles"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   1320
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10398
      _ExtentY        =   661
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Comites y Delegados"
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
      Height          =   420
      Index           =   6
      Left            =   1920
      TabIndex        =   9
      Top             =   360
      Width           =   7485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actividades"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      Left            =   4365
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Reporte"
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
      Index           =   0
      Left            =   4365
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comite"
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
      Left            =   4365
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Promotor"
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
      Index           =   14
      Left            =   4365
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmAF_CD_ReportesComites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim vEstado As Integer, vComite  As String

Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

cboEstado.Enabled = True
cboComite.Enabled = True
cboPromotores.Enabled = True

If Right(Node.Key, 1) = "Z" Then
  lblReporte.Caption = Node.Text
  lblReporte.Tag = fxIndiceCodigo(Node.Key)
  
  Select Case UCase(lblReporte.Tag)
     Case "CM_CMT", "MB_CMT", "ACT_ASG", "CMT_PRM"
       cboTipo.Clear
       cboTipo.AddItem "Detallado"
       cboTipo.Text = "Detallado"
       
       
       dtpInicio.Enabled = True
       dtpCorte.Enabled = True
       
     Case "HIST"
       cboTipo.Clear
       cboTipo.AddItem "Detallado"
       cboTipo.Text = "Detallado"
       cboComite.Enabled = True
       cboPromotores.Enabled = False
       
     Case "LIQ", "LIQ_EST"
       cboTipo.AddItem "Detallado"
       cboTipo.Text = "Detallado"

     Case "ACT"
       cboTipo.Clear
       cboTipo.Text = "Detallado"
     
     Case "CxC", "LG_CXC", "AC_CXC"
       dtpInicio.Enabled = False
       dtpCorte.Enabled = True
       cboTipo.Clear
       cboTipo.AddItem "Resumen"
       cboTipo.AddItem "Detallado"
       cboTipo.Text = "Detallado"
     
     Case "AUX"
       dtpInicio.Enabled = False
       dtpCorte.Enabled = True
       cboTipo.Clear
       cboTipo.AddItem "Resumen"
       cboTipo.AddItem "Detallado"
       cboTipo.Text = "Detallado"
       
  End Select

End If

vError:

End Sub

Private Sub btnInformeEspecial_Click()
Call sbFormsCall("frmAF_CD_InformeEspecial", , , , False, Me)

End Sub

Private Sub btnReporte_Click()
Call sbReporte
End Sub



Private Sub cboComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_COMITE, DESCRIPCION from AFI_CD_COMITES"
       gBusquedas.Filtro = " AND ACTIVO = 1"
       frmBusquedas.Show vbModal
       If gBusquedas.Resultado <> "" Then
         Call sbCboAsignaDato(cboComite, gBusquedas.Resultado2, True, gBusquedas.Resultado)
       End If
End If
End Sub

Private Sub Form_Load()
Dim strSQL As String

 vModulo = 40
 
 
Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture
 
dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
     
strSQL = "select COD_COMITE as 'IdX' , rtrim(DESCRIPCION) as 'ItmX'" _
       & " from AFI_CD_COMITES order by Descripcion"
Call sbCbo_Llena_New(cboComite, strSQL, True, True)
 
strSQL = "select COD_ACTIVIDAD as 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
       & " from  AFI_CD_ACTIVIDADES where ACTIVA = 1 order by Descripcion"
Call sbCbo_Llena_New(cboActividades, strSQL, True, True)
 
strSQL = "select id_promotor as 'IdX', rtrim(nombre) as 'ItmX'" _
       & " from  promotores where Tipo='P' order by Nombre"
Call sbCbo_Llena_New(cboPromotores, strSQL, True, True)
  

cboEstado.Clear
cboEstado.AddItem "Activa"
cboEstado.ItemData(cboEstado.ListCount - 1) = "1"
cboEstado.AddItem "Inactiva"
cboEstado.ItemData(cboEstado.ListCount - 1) = "0"
cboEstado.Text = "Activa"
  
cboTipo.Clear
cboTipo.AddItem "Resumen"
cboTipo.AddItem "Detallado"
cboTipo.Text = "Resumen"
    
  
Call sbRefrescaArbol
 
End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub

Private Sub sbRefrescaArbol()
Dim vNode As Node ', strOpciones  As String
'Dim rs As New ADODB.Recordset, strSQL As String
'Dim vPadre As String


With ArbolExp
  .Nodes.Clear
  Set vNode = .Nodes.Add(, , "Reportes", "Reportes", "imgRoot")
  Call sbCreaNodos("Reportes", "Comites", "imgCRD", False, "0x0CMT")
     Call sbCreaNodos("0x0CMT", "Comites", "imgDetalle", False, "0x0" & "CM_CMT" & "Z")
     Call sbCreaNodos("0x0CMT", "Miembros por Comites", "imgDetalle", False, "0x0" & "MB_CMT" & "Z")
     Call sbCreaNodos("0x0CMT", "Actividades Asignadas", "imgDetalle", False, "0x0" & "ACT_ASG" & "Z")
     Call sbCreaNodos("0x0CMT", "Comites x Promotor", "imgDetalle", False, "0x0" & "CMT_PRM" & "Z")
     Call sbCreaNodos("0x0CMT", "Historial Miembros", "imgDetalle", False, "0x0" & "HIST" & "Z")
  Call sbCreaNodos("Reportes", "Actividades", "imgSGT", False, "0x0ACT")
       Call sbCreaNodos("0x0ACT", "Listado de Actividades", "imgDetalle", False, "0x0" & "ACT" & "Z")
  Call sbCreaNodos("Reportes", "Liquidaciones", "imgCBR", False, "0x0LIQ")
       Call sbCreaNodos("0x0LIQ", "Liquidaciones x Comite", "imgDetalle", False, "0x0" & "LIQ_CMT" & "Z")
       Call sbCreaNodos("0x0LIQ", "Liquidacion x Estado", "imgDetalle", False, "0x0" & "LIQ_EST" & "Z")
  Call sbCreaNodos("Reportes", "Cuentas x Cobrar", "imgCRD", False, "0x0CXC")
       Call sbCreaNodos("0x0CXC", "Listado General", "imgDetalle", False, "0x0" & "LG_CXC" & "Z")
       Call sbCreaNodos("0x0CXC", "Listado Actividades", "imgDetalle", False, "0x0" & "AC_CXC" & "Z")
  Call sbCreaNodos("Reportes", "Auxiliar", "imgCBR", False, "0x0Aux")
       Call sbCreaNodos("0x0Aux", "Auxiliar Contabilidad", "imgDetalle", False, "0x0" & "AUX" & "Z")
     
  .Nodes(1).Expanded = True
End With


End Sub

Private Sub sbReporte()
Dim vTitulo As String, vSubTitulo As String

On Error GoTo vError


If lblReporte.Caption = "" Then
   MsgBox "Debe de seleccionar un reporte!", vbExclamation
   Exit Sub
End If


strSQL = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
    
    vTitulo = UCase(lblReporte.Caption & ", Cuentas por Cobrar : " & cboTipo.Text)
    vSubTitulo = UCase("Comités y Delegados, Fecha de Corte : " & Format(dtpCorte.Value, "dd/MM/yyyy"))
    
  '  cboComite.ItemData (cboComite.ListIndex)
    
  Select Case lblReporte.Caption
    'Comites
       
     Case "Comites"
       .WindowTitle = "Reporte de los Miembros del Comité"
       .ReportFileName = SIFGlobal.fxPathReportes("Comites_Comites.rpt")
       .Formulas(0) = "fxTitulo= 'MIEMBROS DE COMITES'"
     
     Case "Miembros por Comites"
       vTitulo = UCase(lblReporte.Caption & ": " & cboTipo.Text) & " Estado: " & cboEstado.Text
       
       vSubTitulo = "Comités y Delegados, Fechas : " & Format(dtpInicio.Value, "dd/MM/yyyy") & " al " & Format(dtpCorte.Value, "dd/MM/yyyy")
       
       
       
       .WindowTitle = "Reporte de los Miembros del Comité"
       .ReportFileName = SIFGlobal.fxPathReportes("Comites_ListadoMiembrosComite.rpt")
       .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
       .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
       
       If cboComite.Text <> "TODOS" Then
         strSQL = "{AFI_CD_NOMBRAMIENTOS.COD_COMITE} = '" & cboComite.ItemData(cboComite.ListIndex) & "'"
         strSQL = strSQL & " and {AFI_CD_NOMBRAMIENTOS.ACTIVO} = " & cboEstado.ItemData(cboEstado.ListIndex) & ""
       Else
         strSQL = " {AFI_CD_NOMBRAMIENTOS.ACTIVO} = " & cboEstado.ItemData(cboEstado.ListIndex) & ""
       End If
       
         strSQL = strSQL & " AND {AFI_CD_NOMBRAMIENTOS.FECHA_ELECCION} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                   & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
       
       
       .SelectionFormula = strSQL
       
     Case "Actividades Asignadas"
        .WindowTitle = "Reporte Actividad Asignadas al los Comites"
        .ReportFileName = SIFGlobal.fxPathReportes("Comites_ActividadesAsigComite.rpt")
        .Formulas(0) = "fxTitulo= 'ACTIVIDADES ASIGNADAS A LOS COMITES'"
        
        If cboComite.Text <> "TODOS" Then
            .SelectionFormula = "{AFI_CD_COMITES_ACTIVIDADES.COD_COMITE} = '" & cboComite.ItemData(cboComite.ListIndex) & "'"
        End If
     
     Case "Comites x Promotor"
       .WindowTitle = "Comites x Promotor"
       .ReportFileName = SIFGlobal.fxPathReportes("Comites_Comite_Promotor.rpt")
       .StoredProcParam(0) = cboPromotores.ItemData(cboPromotores.ListIndex)
     
     Case "Historial Miembros"
       .WindowTitle = "Historial Miembros"
       .ReportFileName = SIFGlobal.fxPathReportes("Comites_HistorialNombramientos.rpt")
       
       If cboComite.Text <> "TODOS" Then
         strSQL = "{AFI_CD_NOMBRAMIENTOS_H.COD_COMITE} = '" & cboComite.ItemData(cboComite.ListIndex) & "'"
         strSQL = strSQL & " and {AFI_CD_NOMBRAMIENTOS_H.ACTIVO} = " & cboEstado.ItemData(cboEstado.ListIndex)
       Else
         strSQL = " {AFI_CD_NOMBRAMIENTOS_H.ACTIVO} = " & cboEstado.ItemData(cboEstado.ListIndex)
       End If
       
       
       .Formulas(0) = "fxTitulo= 'Historico de Miembros de Cómite'"
       
       .SelectionFormula = strSQL
       
    'Actividades
     Case "Listado de Actividades"
        .WindowTitle = "Reporte de Control de Actividades"
        .ReportFileName = SIFGlobal.fxPathReportes("Comites_actividades.rpt")
        .Formulas(0) = "fxTitulo= 'CONTROL DE ACTIVIDADES'"
        .SelectionFormula = "{AFI_CD_ACTIVIDADES.ACTIVA} = " & cboEstado.ItemData(cboEstado.ListIndex) & ""
        
    'Liquidaciones
     Case "Liquidaciones x Comite"
        .WindowTitle = "Reporte de Control de Actividades"
        .ReportFileName = SIFGlobal.fxPathReportes("Comites_ControlLiquidacionEspecifico.rpt")
        .Formulas(0) = "fxTitulo= 'LIQUIDACIONES REALIZADAS POR LOS COMITES'"
         strSQL = strSQL & " {AFI_CD_CUENTAS.COD_COMITE}='" & cboComite.ItemData(cboComite.ListIndex) & "'"
        .Formulas(4) = "fxFechaInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
        .Formulas(5) = "fxFechaFinal = '" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
        
     Case "Liquidacion x Estado"
        .WindowTitle = "Reporte de Control de Actividades"
        .ReportFileName = SIFGlobal.fxPathReportes("Comites_ControlLiquidacionEspecifico.rpt")
        .Formulas(0) = "fxTitulo= 'LIQUIDACIONES SEGUN SU ESTADO'"
'         strSQL = strSQL & "  {AFI_CD_CUENTAS.ESTADO}='" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
        .Formulas(4) = "fxFechaInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
        .Formulas(5) = "fxFechaFinal = '" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
        
     'Cuentas x Cobrar
     Case "Listado General"
        vTitulo = UCase(lblReporte.Caption & ", Cuentas por Cobrar : " & cboTipo.Text)
        vSubTitulo = UCase("Comités y Delegados, Fecha de Corte : " & Format(dtpCorte.Value, "dd/MM/yyyy"))

        .WindowTitle = "Cuentas por Cobrar Comités y Delegados"
        
        If cboTipo.Text = "Resumen" Then
         .ReportFileName = SIFGlobal.fxPathReportes("Comites_ListadoGeneralResumen.rpt")
        Else
         .ReportFileName = SIFGlobal.fxPathReportes("Comites_ListadoGeneralDetalle.rpt")
        End If
        
        
        .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
        .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
        .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")

     Case "Listado Actividades"
        vTitulo = UCase(lblReporte.Caption & ", Cuentas por Cobrar : " & cboTipo.Text)
        vSubTitulo = UCase("Comités y Delegados, Fecha de Corte : " & Format(dtpCorte.Value, "dd/MM/yyyy"))

        .WindowTitle = "Reporte de Actividades del Comité"
        If cboTipo.Text = "Resumen" Then
         .ReportFileName = SIFGlobal.fxPathReportes("Comites_ListadoActividadesResumen.rpt")
         .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")
        Else
         .ReportFileName = SIFGlobal.fxPathReportes("Comites_ListadoActividadesDetalle.rpt")
         .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")
        End If
       .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
       .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
       
     
     'Auxiliar
     Case "Auxiliar Contabilidad"
       .WindowTitle = "Cuentas por Cobrar Comités y Delegados"
       If cboTipo.Text = "Resumen" Then
        .ReportFileName = SIFGlobal.fxPathReportes("Comites_CuentaContableResumen.rpt")
       Else
        .ReportFileName = SIFGlobal.fxPathReportes("Comites_CuentaContableDetalle.rpt")
       End If
       .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
       .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
       .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")
       
 
    
    End Select
            
    .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "fxUsuario='Usuario: " & glogon.Usuario & "'"
    .Action = 1

End With


Exit Sub

vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

