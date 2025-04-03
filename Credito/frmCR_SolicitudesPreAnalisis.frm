VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_SolicitudesPreAnalisis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reporte de Pre Analisis"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4080
   HelpContextID   =   3022
   Icon            =   "frmCR_SolicitudesPreAnalisis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbPreanalisis 
      Height          =   456
      Left            =   2160
      TabIndex        =   12
      Top             =   2640
      Width           =   1896
      _ExtentX        =   3334
      _ExtentY        =   794
      ButtonWidth     =   2752
      ButtonHeight    =   804
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Pre Analisis"
            Key             =   "preanalisis"
            Object.ToolTipText     =   "Genera Reporte de PreAnalsis"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Cerrar"
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra esta Ventana"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolicitudesPreAnalisis.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolicitudesPreAnalisis.frx":0626
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolicitudesPreAnalisis.frx":0942
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton optPreAnalisis 
      Caption         =   "Pre Analisis por Solicitud"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.OptionButton optPreAnalisis 
      Caption         =   "Pre Analisis por Comité"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Frame fraSolicitud 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
      Begin VB.TextBox txtA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   9
         ToolTipText     =   "Solicitud Final"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDe 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   720
         MaxLength       =   15
         TabIndex        =   7
         ToolTipText     =   "Solicitude de Inicio"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame fraComite 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.ComboBox cboComite 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtActa 
         Height          =   315
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número de Acta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Comité"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCR_SolicitudesPreAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function DatosCompletos() As Boolean
DatosCompletos = False

If optPreAnalisis.Item(0).Value = True Then
  If cboComite.ListIndex <> -1 And Trim(txtActa) = "" Then
     DatosCompletos = True
  ElseIf cboComite.ListIndex <> -1 And Trim(txtActa) <> "" Then
       DatosCompletos = True
  End If
ElseIf optPreAnalisis.Item(1).Value = True Then
  If Trim(txtDe) <> "" And Trim(txtA) <> "" Then
     DatosCompletos = True
  End If
End If
End Function

Private Sub CargacboComite()
Dim rec As New ADODB.Recordset, str As String

On Error GoTo vError

str = "select * from comites"
With rec
.ActiveConnection = glogon.Conection
.CursorType = adOpenStatic
.Source = str
.Open
 Do While Not .EOF And .RecordCount >= 1
 cboComite.AddItem !Descripcion
 cboComite.ItemData(cboComite.NewIndex) = !id_Comite
 .MoveNext
 Loop
 .Close
End With

Set rec = Nothing

Exit Sub
vError:
MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub cboComite_Click()
Dim rs As New ADODB.Recordset
rs.Open "select * from comites where id_comite = " & fxCodigoComite(cboComite.Text), glogon.Conection, adOpenStatic
txtActa = IIf(IsNull(rs!acta), 0, rs!acta)
rs.Close
End Sub


Private Sub Form_Load()
Call CargacboComite
End Sub

Private Sub optPreAnalisis_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0 'Por comité
            fraComite.Enabled = True
            fraSolicitud.Enabled = False
        Case 1 'por Solicitud
            fraComite.Enabled = False
            fraSolicitud.Enabled = True
            txtDe.SetFocus
    End Select
End Sub

Private Sub tlbPreanalisis_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaIngreso As Date

Me.MousePointer = vbHourglass

If DatosCompletos Then
  If optPreAnalisis.Item(0).Value = True Then 'Por Actas
    'Cuando es por acta la membresia no esta bien reflejada
     With frmContenedor.Crt
        .Reset
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .WindowTitle = "Reportes del Módulo de Creditos"
        
        .Connect = glogon.ConectRPT
        
        .ReportFileName = SIFGlobal.fxPathReportes("Credito_PreAnalisis.rpt")
        .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(3) = "TITULO='PREANALSIS POR ACTA'"
        .SelectionFormula = "({REG_CREDITOS.ACTA} = " & txtActa _
               & " AND {REG_CREDITOS.ID_COMITE} = " & fxCodigoComite(cboComite.Text) & ")"
        .PrintReport
     End With
  
  Else
        strSQL = "select id_solicitud,cedula from reg_creditos where estadosol = 'R'" _
               & " and id_solicitud between " & txtDe & " and " & txtA
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            vFechaIngreso = fxMemFechaIngeso(rs!Id_Solicitud)
            With frmContenedor.Crt
              .Reset
              .WindowShowGroupTree = True
              .WindowShowPrintSetupBtn = True
              .WindowShowRefreshBtn = True
              .WindowShowSearchBtn = True
              .WindowState = crptMaximized
              .WindowTitle = "Reportes del Módulo de Creditos"
              
              .Connect = glogon.ConectRPT
              
              .ReportFileName = SIFGlobal.fxPathReportes("Credito_PreAnalisis.rpt")
              .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
              .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
              .Formulas(3) = "TITULO='PREANALSIS POR SOLICITUD'"
              .Formulas(4) = "CATEGORIA='" & fxCalificacionPersona(rs!Cedula) & "'"
              .Formulas(5) = "MEMBRESIA='" & UCase(fxMembresia(vFechaIngreso)) & "'"
              .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & rs!Id_Solicitud
              .PrintReport
            End With
          rs.MoveNext
        Loop
        rs.Close
  End If
     

End If 'DatosCompletos
Me.MousePointer = vbDefault

End Sub

Private Sub txtDe_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtA.SetFocus
End Sub
