VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_ExpedienteGestiones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestiones"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   6255
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtExpediente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Número de Tramite"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Histórico"
      TabPicture(0)   =   "frmFSL_ExpedienteGestiones.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vgGestiones"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Registro"
      TabPicture(1)   =   "frmFSL_ExpedienteGestiones.frx":011B
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line3(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3(7)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAplicar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cboGestiones"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtNotas"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1140
         Width           =   8295
      End
      Begin VB.ComboBox cboGestiones 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8400
         Picture         =   "frmFSL_ExpedienteGestiones.frx":0217
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4260
         Width           =   975
      End
      Begin FPSpreadADO.fpSpread vgGestiones 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   13
         Top             =   600
         Width           =   9975
         _Version        =   524288
         _ExtentX        =   17595
         _ExtentY        =   8493
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         SpreadDesigner  =   "frmFSL_ExpedienteGestiones.frx":0300
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   360
         TabIndex        =   5
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Gestión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   540
         Width           =   855
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   240
         X2              =   9480
         Y1              =   4140
         Y2              =   4140
      End
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteGestiones.frx":092A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteGestiones.frx":718C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteGestiones.frx":7285
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   10320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Cédula"
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
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Expediente"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmFSL_ExpedienteGestiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError

strSQL = "exec spFSL_GestionRegistra " & txtExpediente.Text & ",'" & SIFGlobal.fxSIFCodText(cboGestiones.Text) & "','" _
        & txtNotas.Text & "','" & glogon.Usuario & "'"
glogon.Conection.Execute strSQL

MsgBox "Gestión registrada satisfactoriamente!", vbInformation
Call sbInicializa

Exit Sub

vError:
  MsgBox Err.Description, vbCritical


End Sub

Private Sub Form_Load()

vModulo = 22

txtExpediente.Text = GLOBALES.gTag

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

ssTab.Tab = 0
strSQL = "select rtrim(cod_gestion) + ' - ' + DESCRIPCION as ItmX from FSL_TIPOS_GESTIONES WHERE ACTIVA = 1"
Call sbLlenaCbo(cboGestiones, strSQL, False, False)



strSQL = "select Soc.Nombre,Ex.*" _
       & " from FSL_Expedientes Ex inner join Socios Soc on Ex.cedula = Soc.Cedula" _
       & " Where Ex.Cod_Expediente = " & txtExpediente.Text & " order by registro_fecha desc"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF Or Not rs.BOF Then
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  
  txtNotas.Text = ""
  
 txtEstado.Tag = rs!Estado
  Select Case rs!Estado
   Case "P" 'Pendiente
        txtEstado.Text = "PENDIENTE"
    Case "A" 'Aprobado
        txtEstado.Text = "APROBADO"
    Case "R" 'Rechazado
        txtEstado.Text = "RECHAZADO"
    Case "X" 'Aplicado
        txtEstado.Text = "APLICADO"
  End Select

End If
rs.Close


'Histórico
strSQL = "select Tg.Descripcion, Eg.*" _
       & " from FSL_EXPEDIENTE_GESTIONES Eg inner join FSL_TIPOS_GESTIONES Tg on Eg.COD_GESTION = Tg.COD_GESTION" _
       & " Where Eg.cod_Expediente = " & txtExpediente.Text
rs.Open strSQL, glogon.Conection, adOpenStatic
vgGestiones.MaxRows = 0
Do While Not rs.EOF
  vgGestiones.MaxRows = vgGestiones.MaxRows + 1
  vgGestiones.Row = vgGestiones.MaxRows
  
  vgGestiones.Col = 1
  vgGestiones.Text = rs!Descripcion
  vgGestiones.TextTip = TextTipFixed
  vgGestiones.TextTipDelay = 1000

  vgGestiones.CellNote = "Fecha : " & rs!Registro_Fecha & vbCrLf & "Usuario : " & rs!Registro_Usuario
  vgGestiones.CellTag = CStr(rs!Linea)
    
  vgGestiones.Col = 2
  vgGestiones.Text = rs!notas
      
  vgGestiones.Col = 3
  vgGestiones.Text = rs!Registro_Fecha
      
  vgGestiones.Col = 4
  vgGestiones.Text = rs!Registro_Usuario
 
  vgGestiones.RowHeight(vgGestiones.Row) = vgGestiones.MaxTextRowHeight(vgGestiones.Row)
  
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub
