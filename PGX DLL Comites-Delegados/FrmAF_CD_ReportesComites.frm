VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAF_CD_ReportesComites 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Comites y Delegados"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "FrmAF_CD_ReportesComites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboActividades 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2280
      Width           =   3375
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox cboEstado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ComboBox cboComite 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   3375
   End
   Begin VB.ComboBox cboPromotores 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin MSComctlLib.Toolbar tlbConsulta 
      Height          =   360
      Left            =   7680
      TabIndex        =   0
      Top             =   5040
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      ButtonWidth     =   1799
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "Reporte"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   5760
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
            Picture         =   "FrmAF_CD_ReportesComites.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":359C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":36AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":37D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   4920
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8678
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   120
      Top             =   5760
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
            Picture         =   "FrmAF_CD_ReportesComites.frx":38FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A15F
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A27D
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A3A7
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A4CD
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A5DB
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A6E8
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A801
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":A92F
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ReportesComites.frx":AA3C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   5520
      TabIndex        =   10
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   169738243
      CurrentDate     =   36278
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   5520
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   169738243
      CurrentDate     =   36278
   End
   Begin VB.Label Label1 
      Caption         =   "Actividades"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4365
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4365
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4365
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   9000
      X2              =   4200
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Comites y Delegados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Comite"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4365
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblReporte 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Promotor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   4365
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   8745
      X2              =   8745
      Y1              =   120
      Y2              =   4560
   End
End
Attribute VB_Name = "frmAF_CD_ReportesComites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL, vComite  As String
Dim rs As New ADODB.Recordset
Dim vEstado As Integer

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
  
  cboEstado.Clear
  Call sbCargaEstado
  
  Select Case UCase(lblReporte.Tag)
     Case "CM_CMT", "MB_CMT", "ACT_ASG", "CMT_PRM"
       cboTipo.Clear
       cboTipo.AddItem "Detallado"
       cboTipo.Text = "Detallado"
       dtpInicio.Enabled = False
       dtpCorte.Enabled = False
       
     Case "HIST"
       cboEstado.Enabled = False
       cboEstado.Enabled = False
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

Private Sub Form_Load()
Dim strSQL As String

 vModulo = 23
 
dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
     
strSQL = "select COD_COMITE + ' - ' + rtrim(DESCRIPCION) as ItmX" _
       & " from AFI_CD_COMITES"
Call sbLlenaCbo(cboComite, strSQL)
 
strSQL = "select COD_ACTIVIDAD + ' - ' + rtrim(DESCRIPCION) as ItmX" _
       & " from  AFI_CD_ACTIVIDADES where ACTIVA = 1"
Call sbLlenaCbo(cboActividades, strSQL)
 
strSQL = "select convert(varchar(10),id_promotor) + ' - ' + rtrim(nombre) as ItmX" _
       & " from  promotores where Tipo='P'"
Call sbLlenaCbo(cboPromotores, strSQL)
  

  
Call sbRefrescaArbol
 
End Sub

Sub sbCargaEstado()
    cboEstado.AddItem ("1" & " - " & "Activa")
    cboEstado.AddItem ("0" & " - " & "InActiva")
    cboEstado.Text = "1" & " - " & "Activa"
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
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String
Dim vPadre As String


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
Private Sub tlbConsulta_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, vTitulo As String, vSubTitulo As String
On Error GoTo vError

Select Case UCase(Button.Key)
Case "REPORTE"
 
With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
    
If lblReporte.Caption = Empty Then
   MsgBox "Debe de seleccionar un reporte"
   Exit Sub
Else
    vTitulo = UCase(lblReporte.Caption & ", Cuentas por Cobrar : " & cboTipo.Text)
    vSubTitulo = UCase("Comités y Delegados, Fecha de Corte : " & Format(dtpCorte.Value, "dd/MM/yyyy"))
    Select Case lblReporte.Caption
       
    'Comites
       
     Case "Comites"
       .WindowTitle = "Reporte de los Miembros del Comité"
       .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_Comites.rpt")
       .Formulas(0) = "fxTitulo= 'MIEMBROS DE COMITES'"
     
     Case "Miembros por Comites"
       vTitulo = UCase(lblReporte.Caption & ": " & cboTipo.Text)
       vSubTitulo = UCase("Comités y Delegados, Fecha de Corte : " & Format(dtpCorte.Value, "dd/MM/yyyy"))
       
       .WindowTitle = "Reporte de los Miembros del Comité"
       .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ListadoMiembrosComite.rpt")
       .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
       .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
       If cboComite.Text <> "TODOS" Then
         strSQL = "{AFI_CD_NOMBRAMIENTOS.COD_COMITE} = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
         strSQL = strSQL & " and {AFI_CD_NOMBRAMIENTOS.ACTIVO} = " & SIFGlobal.fxSIFCodText(cboEstado.Text) & ""
       Else
         strSQL = " {AFI_CD_NOMBRAMIENTOS.ACTIVO} = " & SIFGlobal.fxSIFCodText(cboEstado.Text) & ""
       End If
       .SelectionFormula = strSQL
       
     Case "Actividades Asignadas"
        .WindowTitle = "Reporte Actividad Asignadas al los Comites"
        .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ActividadesAsigComite.rpt")
        .Formulas(0) = "fxTitulo= 'ACTIVIDADES ASIGNADAS A LOS COMITES'"
        .SelectionFormula = "{AFI_CD_COMITES_ACTIVIDADES.COD_COMITE} = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
     
     Case "Comites x Promotor"
       .WindowTitle = "Comites x Promotor"
       .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_Comite_Promotor.rpt")
       .StoredProcParam(0) = SIFGlobal.fxSIFCodText(cboPromotores.Text)
     
     Case "Historial Miembros"
       .WindowTitle = "Historial Miembros"
       .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_HistorialNombramientos.rpt")
       .SelectionFormula = "{AFI_CD_NOMBRAMIENTOS_H.COD_COMITE} = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
       .Formulas(0) = "fxTitulo= 'Historico de Miembros de Cómite'"
       
    'Actividades
     Case "Listado de Actividades"
        .WindowTitle = "Reporte de Control de Actividades"
        .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_actividades.rpt")
        .Formulas(0) = "fxTitulo= 'CONTROL DE ACTIVIDADES'"
        .SelectionFormula = "{AFI_CD_ACTIVIDADES.ACTIVA}=" & SIFGlobal.fxSIFCodText(cboEstado.Text) & ""
        
    'Liquidaciones
     Case "Liquidaciones x Comite"
        .WindowTitle = "Reporte de Control de Actividades"
        .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ControlLiquidacionEspecifico.rpt")
        .Formulas(0) = "fxTitulo= 'LIQUIDACIONES REALIZADAS POR LOS COMITES'"
         strSQL = strSQL & " {AFI_CD_CUENTAS.COD_COMITE}='" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
        .Formulas(4) = "fxFechaInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
        .Formulas(5) = "fxFechaFinal = '" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
        
     Case "Liquidacion x Estado"
        .WindowTitle = "Reporte de Control de Actividades"
        .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ControlLiquidacionEspecifico.rpt")
        .Formulas(0) = "fxTitulo= 'LIQUIDACIONES SEGUN SU ESTADO'"
         strSQL = strSQL & "  {AFI_CD_CUENTAS.ESTADO}='" & SIFGlobal.fxSIFCodText(cboEstado.Text) & "'"
        .Formulas(4) = "fxFechaInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
        .Formulas(5) = "fxFechaFinal = '" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
        
     'Cuentas x Cobrar
     Case "Listado General"
        vTitulo = UCase(lblReporte.Caption & ", Cuentas por Cobrar : " & cboTipo.Text)
        vSubTitulo = UCase("Comités y Delegados, Fecha de Corte : " & Format(dtpCorte.Value, "dd/MM/yyyy"))

        .WindowTitle = "Cuentas por Cobrar Comités y Delegados"
        If cboTipo.Text = "Resumen" Then
         .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ListadoGeneralResumen.rpt")
        Else
         .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ListadoGeneralDetalle.rpt")
        End If
        .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
        .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
        .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")

     Case "Listado Actividades"
        vTitulo = UCase(lblReporte.Caption & ", Cuentas por Cobrar : " & cboTipo.Text)
        vSubTitulo = UCase("Comités y Delegados, Fecha de Corte : " & Format(dtpCorte.Value, "dd/MM/yyyy"))

        .WindowTitle = "Reporte de Actividades del Comité"
        If cboTipo.Text = "Resumen" Then
         .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ListadoActividadesResumen.rpt")
         .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")
        Else
         .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ListadoActividadesDetalle.rpt")
         .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")
        End If
       .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
       .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
       
     
     'Auxiliar
     Case "Auxiliar Contabilidad"
       .WindowTitle = "Cuentas por Cobrar Comités y Delegados"
       If cboTipo.Text = "Resumen" Then
        .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_CuentaContableResumen.rpt")
       Else
        .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_CuentaContableDetalle.rpt")
       End If
       .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
       .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
       .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")
       
 
    
    End Select
            
.Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
.Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
.Formulas(3) = "fxUsuario='Usuario: " & glogon.Usuario & "'"
.Action = 1
'.PrintReport

End If
End With
End Select

Exit Sub

vError:
      MsgBox Err.Description

End Sub






