VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmSYS_Load_File 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga de Archivo"
   ClientHeight    =   7740
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkProcesar 
      Caption         =   "Cargar y Subir"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10560
      TabIndex        =   25
      Top             =   6720
      Width           =   1332
   End
   Begin MSComctlLib.ProgressBar ProgressBarX 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   24
      Top             =   7620
      Width           =   12624
      _ExtentX        =   22278
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4572
      Left            =   10440
      TabIndex        =   9
      Top             =   1920
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   8064
      _StockProps     =   79
      Caption         =   "Referencias a Subir?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BorderStyle     =   1
      Begin VB.ComboBox cboReferencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   5
         ItemData        =   "frmSYS_Load_File.frx":0000
         Left            =   720
         List            =   "frmSYS_Load_File.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2400
         Width           =   1092
      End
      Begin VB.ComboBox cboReferencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   4
         ItemData        =   "frmSYS_Load_File.frx":0076
         Left            =   720
         List            =   "frmSYS_Load_File.frx":0098
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2040
         Width           =   1092
      End
      Begin VB.ComboBox cboReferencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   3
         ItemData        =   "frmSYS_Load_File.frx":00EC
         Left            =   720
         List            =   "frmSYS_Load_File.frx":010E
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   1092
      End
      Begin VB.ComboBox cboReferencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   2
         ItemData        =   "frmSYS_Load_File.frx":0162
         Left            =   720
         List            =   "frmSYS_Load_File.frx":0184
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   1092
      End
      Begin VB.ComboBox cboReferencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   1
         ItemData        =   "frmSYS_Load_File.frx":01D8
         Left            =   720
         List            =   "frmSYS_Load_File.frx":01FA
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   1092
      End
      Begin VB.ComboBox cboReferencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   0
         ItemData        =   "frmSYS_Load_File.frx":024E
         Left            =   720
         List            =   "frmSYS_Load_File.frx":0270
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label lblLoading 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Líneas Cargadas?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "REF 6"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "REF 5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "REF 4"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "REF 3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "REF 2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "REF 1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   732
      End
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   6840
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No Procesar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   372
      Left            =   10560
      TabIndex        =   5
      Top             =   7080
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Subir Datos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
   End
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   8895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":02C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":6B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":D388
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":13BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1A44C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1AC0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1B2D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1BCC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1C684
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1D0A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1D87A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1E237
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSYS_Load_File.frx":1E92E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   690
      Left            =   10320
      TabIndex        =   2
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1217
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buscar"
            Object.ToolTipText     =   "Buscar archivos"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cargar"
            Object.ToolTipText     =   "Cargar información"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4692
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   9972
      _Version        =   524288
      _ExtentX        =   17590
      _ExtentY        =   8276
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
      MaxCols         =   10
      SpreadDesigner  =   "frmSYS_Load_File.frx":1F0EA
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   6840
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Distribución Geografica"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   7200
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Padron Nacional"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CommonDialog cmd 
      Left            =   240
      Top             =   1440
      _Version        =   1310723
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.TaskDialog TaskDialog1 
      Left            =   1080
      Top             =   2880
      _Version        =   1310723
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carga de Archivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4452
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmSYS_Load_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim pCon As New ADODB.Connection, pFile As Long, pLinea As Long
'
'Private Sub btnCargar_Click()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim i As Long, pPais As String
'Dim Ref(5) As String, pArchivo As String
'
'On Error GoTo vError
'
'Me.MousePointer = vbHourglass
'
'pPais = "CRC"
'If chkProcesar.Value = vbUnchecked _
' Or chkProcesar.Value = vbChecked And pFile = 0 Then
'    pArchivo = Dir(txtArchivo.Text, vbArchive)
'    strSQL = "exec spSys_Load_File '" & pArchivo & "','" & glogon.Usuario & "'"
'    rs.Open strSQL, glogon.BaseCon, adOpenStatic
'       pFile = rs!File_ID
'    rs.Close
'    pLinea = 1
'End If
'
'strSQL = ""
'
'With vGrid
'
'ProgressBarX.Max = .MaxRows
'For i = 1 To .MaxRows
'   .Row = i
'   .Col = Val(Right(cboReferencia.Item(0).Text, 2))
'   Ref(0) = Trim(.Text)
'   .Col = Val(Right(cboReferencia.Item(1).Text, 2))
'   Ref(1) = Trim(.Text)
'   .Col = Val(Right(cboReferencia.Item(2).Text, 2))
'   Ref(2) = Trim(.Text)
'   .Col = Val(Right(cboReferencia.Item(3).Text, 2))
'   Ref(3) = Trim(.Text)
'   .Col = Val(Right(cboReferencia.Item(4).Text, 2))
'   Ref(4) = Trim(.Text)
'   .Col = Val(Right(cboReferencia.Item(5).Text, 2))
'   Ref(5) = Trim(.Text)
'
'   If chkProcesar.Value = vbChecked Then
'      pLinea = pLinea + 1
'        strSQL = strSQL & Space(10) & "exec spSys_Load_File_Detalle " & pFile & "," & pLinea & ",'" & Ref(0) & "','" & Ref(1) _
'               & "','" & Ref(2) & "','" & Ref(3) & "','" & Ref(4) & "','" & Ref(5) & "'"
'   Else
'        strSQL = strSQL & Space(10) & "exec spSys_Load_File_Detalle " & pFile & "," & i & ",'" & Ref(0) & "','" & Ref(1) _
'               & "','" & Ref(2) & "','" & Ref(3) & "','" & Ref(4) & "','" & Ref(5) & "'"
'   End If
'   If Len(strSQL) > 20000 Then
'      glogon.BaseCon.Execute strSQL
'      strSQL = ""
'   End If
'
'   ProgressBarX.Value = i
'
'Next i
'
''Ultimo Lote
'If Len(strSQL) > 0 Then
'   glogon.BaseCon.Execute strSQL
'End If
'
'End With
'
'If chkProcesar.Value = vbChecked Then
'   vGrid.MaxRows = 0
'Else
'    Me.MousePointer = vbDefault
'    MsgBox "Información Subida a la Base de Datos!", vbInformation
'End If
'Exit Sub
'
'vError:
'  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler (Err.Description), vbCritical
'
'End Sub
'
'Private Sub Form_Load()
'Dim strSQL As String
'
'vModulo = 13
'vGrid.AppearanceStyle = fxGridStyle
'
'Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture
'
'Call Formularios(Me)
'Call RefrescaTags(Me)
'End Sub
'
'Private Sub optX_Click(Index As Integer)
'Select Case Index
'  Case 0 'Ninguno
'    cboReferencia.Item(0).Text = "REF_01"
'    cboReferencia.Item(1).Text = "REF_02"
'    cboReferencia.Item(2).Text = "REF_03"
'    cboReferencia.Item(3).Text = "REF_04"
'    cboReferencia.Item(4).Text = "REF_05"
'    cboReferencia.Item(5).Text = "REF_06"
'  Case 1 'Distribucion
'    cboReferencia.Item(0).Text = "REF_01"
'    cboReferencia.Item(1).Text = "REF_02"
'    cboReferencia.Item(2).Text = "REF_03"
'    cboReferencia.Item(3).Text = "REF_04"
'    cboReferencia.Item(4).Text = "REF_05"
'    cboReferencia.Item(5).Text = "REF_06"
'  Case 2 'Padron
'    cboReferencia.Item(0).Text = "REF_01"
'    cboReferencia.Item(1).Text = "REF_02"
'    cboReferencia.Item(2).Text = "REF_04"
'    cboReferencia.Item(3).Text = "REF_06"
'    cboReferencia.Item(4).Text = "REF_07"
'    cboReferencia.Item(5).Text = "REF_08"
'
'End Select
'
'End Sub
'
'Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
'Select Case Button.Key
'  Case "buscar"
'
'        txtArchivo.Text = ""
'
'        With cmd
'                .InitDir = "C:\"
'                .DialogTitle = "Localice Archivo de Trama [Texto]..."
'                .Filter = "*.txt"
'                .ShowOpen
'
'                If .FileName = "" Then
'                  MsgBox "Archivo no válido...", vbExclamation
'                  Exit Sub
'                End If
'
'                If UCase(Right(.FileName, 3)) <> "TXT" Then
'                  MsgBox "La Extensión del Archivo no es válido...", vbExclamation
'                  Exit Sub
'                End If
'
'         txtArchivo.Text = .FileName
'
'        End With
'
'  Case "cargar"
'    Call sbArchivo_Lee
'
'End Select
'
'
'End Sub
'
'
'
'Private Sub sbArchivo_Lee()
'Dim strCadena As String
'Dim fn, Columna As Integer, Campos() As String
'
'
'On Error GoTo vError
'
'If txtArchivo.Text = "" Then
'   MsgBox "Seleccione un archivo a procesar...", vbExclamation
'   Exit Sub
'End If
'
'Me.MousePointer = vbHourglass
'
'vGrid.MaxRows = 0
'
'fn = FreeFile
'Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
' Do While Not EOF(fn)
'   Line Input #fn, strCadena
'
'
'   If vGrid.MaxRows = 200000 And chkProcesar.Value = vbChecked Then
'      Call btnCargar_Click
'   End If
'
'   vGrid.MaxRows = vGrid.MaxRows + 1
'   vGrid.Row = vGrid.MaxRows
'
'   lblLoading.Caption = vGrid.MaxRows
'   DoEvents
'
'   vGrid.Col = 6
'   vGrid.Text = strCadena
'
'    'Lee todas las columnas
'    Campos = Split(strCadena, ",")
'    For Columna = 0 To UBound(Campos)
'      If Columna <= 10 Then '10 Referencias
'         vGrid.Col = Columna + 1
'         vGrid.Text = Campos(Columna)
'      End If
'
'    Next Columna
' Loop
'Close #fn
'
'
'   If vGrid.MaxRows > 0 And chkProcesar.Value = vbChecked Then
'      Call btnCargar_Click
'   End If
'
'MsgBox "Información Cargada Satisfactoriamente!", vbInformation
'
'lblLoading.Caption = Format(vGrid.MaxRows, "###,###,###,###,###")
'
'Me.MousePointer = vbDefault
'
'Exit Sub
'
'vError:
'    Me.MousePointer = vbDefault
'    MsgBox fxSys_Error_Handler (Err.Description), vbCritical
'End Sub
'
'
'Private Sub sbArchivo_Sube()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim strCadena As String, curMonto As Currency
'Dim fn, Casos(4) As Long, Campos() As String
'
'
'Dim pTipoPoliza As String, pPoliza As String, pLinea As Long, pMoneda As String, pAseguradora As String
'Dim pTipoId As String, pCedula As String, pNombre As String, pNumCta As Integer, pMonto As Currency
'Dim pExiste As Integer, pTarjetaNum As String, pTarjetaVence As String, pComision As Currency
'Dim pMedioPago As String, pFechaInicio As Date, pFechaCorte As Date, pTempo As String
'
'On Error GoTo vError
'
'If txtArchivo.Text = "" Then
'   MsgBox "Seleccione un archivo a procesar...", vbExclamation
'   Exit Sub
'End If
'
'Me.MousePointer = vbHourglass
'
'vGrid.MaxRows = 0
'
'curMonto = 0
'
'Casos(0) = 0 'Total
'Casos(1) = 0 'Existe
'Casos(2) = 0 'No Existe
'Casos(3) = 0 'Cambios
'
'pLinea = 0
'strSQL = ""
''pAseguradora = SIFGlobal.fxCodText(cboAseguradora.Text)
'
'fn = FreeFile
'Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
' Do While Not EOF(fn)
'   Line Input #fn, strCadena
'   Campos = Split(strCadena, ",")
'
'' Script para Leer todas las columnas
''    For Columna = 0 To UBound(Campos)
''      x  = Campos(Columna)
''    Next Columna
'
'   pLinea = pLinea + 1
'   pNumCta = Campos(0)          'fxTramaCampo(strCadena, 1)
'   pMoneda = Campos(1)          'fxTramaCampo(strCadena, 2)
'   pPoliza = Campos(2)          'fxTramaCampo(strCadena, 3)
'   pTarjetaNum = Campos(4)      'fxTramaCampo(strCadena, 5)
'   pTarjetaVence = Campos(5)    'fxTramaCampo(strCadena, 6)
'   pMedioPago = Campos(6)       'fxTramaCampo(strCadena, 7)
'   pTempo = Campos(8)           'fxTramaCampo(strCadena, 9)
'   If IsDate(pTempo) Then
'       pFechaInicio = pTempo
'   Else
'       pFechaInicio = dtpCuota.Value
'   End If
'   pTempo = Campos(9)           'fxTramaCampo(strCadena, 9)
'   If IsDate(pTempo) Then
'       pFechaCorte = pTempo
'   Else
'       pFechaCorte = dtpVence.Value
'   End If
'
'   pMonto = Campos(10)          'fxTramaCampo(strCadena, 11)
'   pComision = Campos(12)       'fxTramaCampo(strCadena, 13)
'   pTipoId = Trim(Campos(14))         'fxTramaCampo(strCadena, 14)
'   pCedula = Campos(13)         'fxTramaCampo(strCadena, 15)
'   pNombre = UCase(Campos(15))  'fxTramaCampo(strCadena, 16)
'
'
'   strSQL = strSQL & Space(10) & "insert seguros_Tramas(cod_aseguradora,Trama_Id, Cod_Linea,Num_Cuota, Num_Poliza, Monto, Comision_Neta" _
'                     & ",Tipo_Id, Cedula, Nombre, Fecha_Cuota, Fecha_Vence, Moneda" _
'                     & ",Medio_pago, Tarjeta_Numero, Tarjeta_Vence, Trama_Original, Registro_Fecha, Registro_Usuario)" _
'                     & " Values('" & pAseguradora & "','" & txtTramaId.Text & "'," & pLinea & "," & pNumCta & ",'" & pPoliza & "'," & pMonto & "," & pComision _
'                     & ",'" & pTipoId & "','" & pCedula & "','" & pNombre & "','" & Format(pFechaInicio, "yyyy/mm/dd") & "','" & Format(pFechaCorte, "yyyy/mm/dd") _
'                     & "','" & pMoneda & "','" & pMedioPago & "','" & pTarjetaNum & "','" & pTarjetaVence & "','" & strCadena _
'                     & "', getdate(),'" & glogon.Usuario & "')"
'
'   If Len(strSQL) > 20000 Then
'      Call ConectionExecute(strSQL)
'      strSQL = ""
'   End If
'
' Loop
'Close #fn
'
''Cadena Final
'If Len(strSQL) > 0 Then
'    Call ConectionExecute(strSQL)
'    strSQL = ""
'End If
'
''Verifica y Devuelve Resultado
'Call sbTrama_Verificada_Load(pAseguradora, txtTramaId.Text, 1, 1)
'
'Me.MousePointer = vbDefault
'
'Exit Sub
'
'vError:
'    Me.MousePointer = vbDefault
'    MsgBox fxSys_Error_Handler (Err.Description), vbCritical
'End Sub
'
