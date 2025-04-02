VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_ControlCasosEjecutivos 
   Caption         =   "Créditos: Control de Casos por Ejecutivos"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBarX 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   14
      Top             =   8115
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtFiltroEjecutivo 
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
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
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
            Picture         =   "frmCR_ControlCasosEjecutivos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ControlCasosEjecutivos.frx":00F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ControlCasosEjecutivos.frx":0223
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   7335
      Left            =   4680
      TabIndex        =   0
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Asignados"
      TabPicture(0)   =   "frmCR_ControlCasosEjecutivos.frx":032C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Nuevos"
      TabPicture(1)   =   "frmCR_ControlCasosEjecutivos.frx":0425
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboAsignadoA"
      Tab(1).Control(1)=   "dtpInicio"
      Tab(1).Control(2)=   "cboRecibidosPor"
      Tab(1).Control(3)=   "chkCasos"
      Tab(1).Control(4)=   "cboOficina"
      Tab(1).Control(5)=   "dtpCorte"
      Tab(1).Control(6)=   "vGridAsg"
      Tab(1).Control(7)=   "Line1"
      Tab(1).Control(8)=   "Label2(3)"
      Tab(1).Control(9)=   "Label2(2)"
      Tab(1).Control(10)=   "Label2(1)"
      Tab(1).Control(11)=   "Label2(0)"
      Tab(1).ControlCount=   12
      Begin VB.ComboBox cboAsignadoA 
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
         Left            =   -73320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1200
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   -67560
         TabIndex        =   10
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   241041411
         CurrentDate     =   41785
      End
      Begin VB.ComboBox cboRecibidosPor 
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
         Left            =   -73320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   4575
      End
      Begin VB.CheckBox chkCasos 
         Caption         =   "Marcar"
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
         Left            =   -74760
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cboOficina 
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
         Left            =   -73320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   315
         Left            =   -67560
         TabIndex        =   11
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   241041411
         CurrentDate     =   41785
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6732
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   8532
         _Version        =   524288
         _ExtentX        =   15050
         _ExtentY        =   11875
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
         MaxCols         =   14
         SpreadDesigner  =   "frmCR_ControlCasosEjecutivos.frx":052C
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridAsg 
         Height          =   5292
         Left            =   -74880
         TabIndex        =   16
         Top             =   1920
         Width           =   8412
         _Version        =   524288
         _ExtentX        =   14838
         _ExtentY        =   9335
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
         MaxCols         =   14
         SpreadDesigner  =   "frmCR_ControlCasosEjecutivos.frx":0E1C
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74760
         X2              =   -66480
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Asignado a:"
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
         Left            =   -74760
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Formalizados entre"
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
         Index           =   2
         Left            =   -68640
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Recibidos por:"
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
         Left            =   -74760
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Oficina"
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
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   12938
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.Label lblEjecutivo 
      Alignment       =   1  'Right Justify
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmCR_ControlCasosEjecutivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Load()
vModulo = 3

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub




Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

ssTab.Tab = 0


lblEjecutivo.Tag = ""
lblEjecutivo.Caption = ""

vPaso = True


strSQL = "select ID_PROMOTOR, NOMBRE " _
       & "  From PROMOTORES" _
       & " Where Estado = 1 and Nombre like '%" & Trim(txtFiltroEjecutivo.Text) & "%'" _
       & " ORDER BY NOMBRE"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Nombre)
      itmX.Tag = rs!ID_PROMOTOR
  rs.MoveNext
Loop
rs.Close

vPaso = False


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Resize()
On Error Resume Next

lsw.Height = Me.Height - (lsw.top + ProgressBarX.Height + 350)
ssTab.Height = lsw.Height

ssTab.Left = lsw.Left
ssTab.Width = Me.Width - 300


End Sub

Private Sub lsw_Click()
If vPaso Then Exit Sub
If lsw.ListItems.Count = 0 Then Exit Sub

lblEjecutivo.Tag = lsw.SelectedItem.Tag
lblEjecutivo.Caption = lsw.SelectedItem.Text

ssTab.SetFocus

End Sub

Private Sub txtFiltroEjecutivo_GotFocus()
On Error Resume Next
ssTab.Left = lsw.Left + lsw.Width + 60
ssTab.Width = Me.Width - (ssTab.Left + 300)

End Sub

Private Sub txtFiltroEjecutivo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
   Call sbInicializa
End If

End Sub

Private Sub txtFiltroEjecutivo_LostFocus()
Call Form_Resize
End Sub
