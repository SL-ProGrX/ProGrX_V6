VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCR_PolizasRemesasPago 
   Caption         =   "Control de Polizas: Proceso de Pagos"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   11370
   Begin VB.Data DaoControl 
      Caption         =   "DaoControl"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Tag             =   "u"
      Top             =   240
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabCaption(0)   =   "Recepción"
      TabPicture(0)   =   "frmCR_PolizasRemesasPago.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpVence"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "vGrid"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tlbProceso"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tlbX"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dtpCuota"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtMonto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCasos"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtExiste"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNoExiste"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCambio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtArchivo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Envío al INS"
      TabPicture(1)   =   "frmCR_PolizasRemesasPago.frx":00FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboRemesa"
      Tab(1).Control(1)=   "lsw"
      Tab(1).Control(2)=   "tlbTrama"
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).Control(4)=   "Line1(0)"
      Tab(1).Control(5)=   "Label1(4)"
      Tab(1).Control(6)=   "Line1(1)"
      Tab(1).ControlCount=   7
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
         TabIndex        =   9
         Top             =   600
         Width           =   8895
      End
      Begin VB.TextBox txtCambio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtNoExiste 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtExiste 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtCasos 
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
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Monto"
         Top             =   6000
         Width           =   1575
      End
      Begin VB.ComboBox cboRemesa 
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
         Left            =   -72840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   7575
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   4335
         Left            =   -72840
         TabIndex        =   1
         Top             =   1200
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7646
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Casos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   4304
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   1440
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
         Format          =   171245571
         CurrentDate     =   41106
      End
      Begin MSComctlLib.Toolbar tlbX 
         Height          =   660
         Left            =   10320
         TabIndex        =   10
         Top             =   600
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   1164
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
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cargar"
               Object.ToolTipText     =   "Cargar información"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbProceso 
         Height          =   330
         Left            =   7800
         TabIndex        =   11
         Top             =   5880
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "aplicar"
               Object.ToolTipText     =   "Aplicar Archivo"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "cancelar"
               Object.ToolTipText     =   "cancelar operacion"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3855
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   6800
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
         MaxCols         =   495
         SpreadDesigner  =   "frmCR_PolizasRemesasPago.frx":01FB
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpVence 
         Height          =   315
         Left            =   9000
         TabIndex        =   13
         Top             =   1440
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
         Format          =   171245571
         CurrentDate     =   41106
      End
      Begin MSComctlLib.Toolbar tlbTrama 
         Height          =   360
         Left            =   -68280
         TabIndex        =   14
         Top             =   6000
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   635
         ButtonWidth     =   5054
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Procesar Tramas Seleccionadas"
               Key             =   "Procesar"
               ImageIndex      =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
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
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cambios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   23
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Existe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   22
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Existe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   21
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Casos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   20
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de la Cuota..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha vencimiento de Pago..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6360
         TabIndex        =   17
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Remesa de Pago"
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
         Index           =   3
         Left            =   -74640
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -74760
         X2              =   -64200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Tramas Disponibles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   -74640
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74760
         X2              =   -64200
         Y1              =   5640
         Y2              =   5640
      End
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   5040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   120
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
            Picture         =   "frmCR_PolizasRemesasPago.frx":086A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasRemesasPago.frx":70CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasRemesasPago.frx":D92E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasRemesasPago.frx":14190
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCR_PolizasRemesasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Function fxRemueveCeroIzq(vCadena As String) As String
vPaso = True

Do While vPaso
   If Mid(vCadena, 1, 1) = "0" Then
      vCadena = Mid(vCadena, 2, Len(vCadena))
   Else
      vPaso = False
   End If
Loop

fxRemueveCeroIzq = vCadena

End Function


Private Sub sbCargaTrama()
Dim strCadena As String, curMonto As Currency
Dim fn, Casos(4) As Long, pMonto As Currency
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

curMonto = 0

Casos(0) = 0 'Total
Casos(1) = 0 'Existe
Casos(2) = 0 'No Existe
Casos(3) = 0 'Cambios



fn = FreeFile
Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
 Do While Not EOF(fn)
   Input #fn, strCadena
   
   If Len(strCadena) >= 66 Then 'Largo de la Trama
            With vGrid
               .MaxRows = .MaxRows + 1
               .Row = .MaxRows
               
               .Col = 1 'Cedula
               .Text = fxRemueveCeroIzq(Mid(strCadena, 28, 20))
               
               .Col = 2 'Nombre
               .Text = fxNombre(fxRemueveCeroIzq(Mid(strCadena, 28, 20)))
               
               .Col = 3 'No. Cuota
               .Text = fxRemueveCeroIzq(Mid(strCadena, 48, 4))
               
               pMonto = CCur(fxRemueveCeroIzq(Mid(strCadena, 52, 15))) / 100
               curMonto = curMonto + pMonto
             
               .Col = 4 'Monto
               .Text = pMonto
               
               .Col = 5 'No. Póliza
               .Text = Mid(strCadena, 5, 20)
               
                strSQL = "select * from Ins_Polizas where num_poliza = '" & .Text & "'"
                
               
               .Col = 6 'Tipo de Seguro
               .Text = Mid(strCadena, 1, 3)
               
               
                rs.Open strSQL, glogon.Conection, adOpenStatic
                If Not rs.EOF And Not rs.BOF Then
                  .Col = 7
                  .Value = vbChecked
                  Casos(1) = Casos(1) + 1 'Existe
                  
                  If pMonto <> rs!Cuota Then Casos(3) = Casos(3) + 1
                Else
                  .Col = 7
                  .Value = vbUnchecked
                  Casos(2) = Casos(2) + 1 'No Existe
                End If
                rs.Close
            
           End With
     End If 'Len(strCadena) >= 66
 
 Loop
Close #fn
        


'Totales
txtMonto.Text = Format(curMonto, "Standard")
txtCasos.Text = vGrid.MaxRows

txtExiste.Text = Casos(1)
txtNoExiste.Text = Casos(2)
txtCambio.Text = Casos(3)


Me.MousePointer = vbDefault

If Casos(2) = 0 Then
    MsgBox "Información Cargada Satisfactoriamente", vbInformation
Else
    MsgBox "Información Cargada Pero Existen varios casos que no estan registrados!", vbInformation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtExiste.Text = 0
    txtNoExiste.Text = 0
    txtCambio.Text = 0
End Sub

Private Sub cboRemesa_Click()
If vPaso Or cboRemesa.ListCount = 0 Then Exit Sub
Call sbLswTramas
End Sub


Private Sub dtpCuota_Change()
Dim vFecha As Date

On Error GoTo vError
vFecha = DateAdd("m", 1, dtpCuota.Value)
vFecha = CDate(Year(vFecha) & "/" & Format(Month(vFecha), "00") & "/30")

dtpVence.Value = vFecha

Exit Sub

vError:
vFecha = CDate(Year(vFecha) & "/" & Format(Month(vFecha), "00") & "/28")

dtpVence.Value = vFecha
 

End Sub

Private Sub Form_Activate()

vModulo = 11

End Sub

Private Sub Form_Load()


vModulo = 11

vGrid.MaxCols = 7
vGrid.MaxRows = 0

ssTab.Tab = 0
dtpCuota.Value = fxFechaServidor
Call dtpCuota_Change

End Sub


Private Sub sbProcesar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCedula As String, pNombre As String, pCtaNum As Integer, pMonto As Currency, pPolizaNum As String, pPolizaSeg As String
Dim lng As Long, vProcesados As Long


On Error GoTo vError

Me.MousePointer = vbHourglass

vProcesados = 0

With vGrid
    For lng = 1 To .MaxRows
       .Row = lng
       .Col = 7
       If .Value = vbChecked Then 'El Caso Existe
            .Col = 1
            pCedula = Trim(.Text)
            .Col = 2
            pNombre = Trim(.Text)
            .Col = 3
            pCtaNum = .Text
            .Col = 4
            pMonto = CCur(.Text)
            .Col = 5
            pPolizaNum = Trim(.Text)
            .Col = 6
            pPolizaSeg = Trim(.Text)
               
            vProcesados = vProcesados + 1
               
            'Actualiza datos de la póliza
            strSQL = "update Ins_Polizas set cuota = " & pMonto & ",num_cuota = " & pCtaNum _
                   & ", Tipo_Seguro = '" & pPolizaSeg & "'  where num_poliza = '" & pPolizaNum & "'"
            glogon.Conection.Execute strSQL
            
            'Registra Cuota al Cobro
            strSQL = "insert ins_pagos(num_poliza,num_cuota,monto,monto_neto,monto_pago,fecha_vence,fecha_cuota,registro_fecha,registro_usuario)" _
                   & " values('" & pPolizaNum & "'," & pCtaNum & "," & pMonto & "," & pMonto & ",0,'" & Format(dtpVence.Value, "yyyy/mm/dd") _
                   & "','" & Format(dtpCuota.Value, "yyyy/mm/dd") & "',getdate(),'" & glogon.Usuario & "')"
            glogon.Conection.Execute strSQL
               
       End If '.Value = vbChecked
    
    Next lng

End With

Me.MousePointer = vbDefault

MsgBox "Carga de Pólizas para el cobro...aplicadas Satisfactoriamente... Registros Procesados :" & vProcesados, vbInformation

txtArchivo.Text = ""
vGrid.MaxRows = 0

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub sbLswTramas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass
 
'Falta el filtro de la remesa

strSQL = "select FECHA_VENCE, COUNT(*) as 'Casos', SUM(MONTO_PAGO) as 'Monto'" _
       & " From INS_PAGOS" _
       & " Where NUM_CUOTA > 0 and cod_remesa = " & cboRemesa.ItemData(cboRemesa.ListIndex) _
       & " group by FECHA_VENCE"
lsw.ListItems.Clear
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Format(rs!Fecha_Vence, "dd/mm/yyyy"))
     itmX.SubItems(1) = rs!Casos
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.Checked = True

 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
  
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset


If ssTab.Tab = 1 Then
    vPaso = True
    
    cboRemesa.Clear
    
    strSQL = "select Top 100 * from INS_REMESAS where Tipo = 'A' and estado in('C','T') order by fecha desc"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rs.EOF
      cboRemesa.AddItem (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!Fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_CORTE, "dd/mm/yyyy"))
      cboRemesa.ItemData(cboRemesa.NewIndex) = rs!cod_remesa
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboRemesa.Text = (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!Fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_CORTE, "dd/mm/yyyy"))
    End If
    rs.Close
    vPaso = False


   Call sbLswTramas
End If

End Sub


Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "aplicar"
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen casos a procesar...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
  
  Case "cancelar"
    vGrid.MaxRows = 0
    txtArchivo.Text = ""

End Select
End Sub



Private Sub sbCreaTrama(vFecha As Date)
Dim i As Long, vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String
Dim fnFile

Dim strSQL As String, rs As New ADODB.Recordset, lngMonto As Long

fnFile = FreeFile


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\INS"
vRuta = SIFGlobal.DirectorioDeResultados & "\INS"

vArchivo = "CC" & Format(vFecha, "yyyymmdd") & ".txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError

Open vTempo For Output As #fnFile  ' Create file name.

strSQL = "select Pol.TIPO_SEGURO,4 as 'Tipo_Mov', Pol.NUM_POLIZA,Pag.FECHA_VENCE,Pol.TIPO_CUENTA" _
       & ",Pol.CEDULA,Pag.NUM_CUOTA,Pag.MONTO_PAGO, Pag.Monto, case when Pag.MONTO_PAGO = 0 then '001' else '000' end as 'Motivo' " _
       & " from INS_POLIZAS Pol inner join INS_PAGOS Pag on Pol.NUM_POLIZA = Pag.NUM_POLIZA" _
       & " where Pag.NUM_CUOTA > 0 and Pag.FECHA_VENCE = '" & Format(vFecha, "yyyy/mm/dd") & "'" _
       & " and Pag.Cod_Remesa = " & cboRemesa.ItemData(cboRemesa.ListIndex)
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 If rs!Motivo = "001" Then
    lngMonto = 0
 Else
    lngMonto = CLng(rs!Monto * 100)
 End If
 vCadena = Trim(rs!Tipo_Seguro)
 vCadena = vCadena & rs!Tipo_Mov
 vCadena = vCadena & SIFGlobal.fxSIFRelleno(rs!Num_Poliza, "D", " ", 20)
 vCadena = vCadena & Format(vFecha, "yyyymmdd")
 vCadena = vCadena & Trim(rs!Tipo_Cuenta)
 vCadena = vCadena & SIFGlobal.fxSIFRelleno(rs!Cedula, "I", "0", 20)
 vCadena = vCadena & SIFGlobal.fxSIFRelleno(rs!num_cuota, "I", "0", 4)
 vCadena = vCadena & SIFGlobal.fxSIFRelleno(CStr(lngMonto), "I", "0", 15)
 vCadena = vCadena & Trim(rs!Motivo)
 
 Print #fnFile, vCadena

 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
End Sub


Private Sub tlbTrama_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

On Error GoTo vError

Select Case Button.Key

  Case "Procesar"
        For i = 1 To lsw.ListItems.Count
           If lsw.ListItems.Item(i).Checked Then
        
             Call sbCreaTrama(lsw.ListItems.Item(i).Text)

           End If
        Next i
 
 Case Else

End Select

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "buscar"
        
        txtArchivo.Text = ""
        
        With Cmd
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Trama [Texto]..."
                .Filter = "*.txt"
                .ShowOpen
                
                If .FileName = "" Then
                  MsgBox "Archivo no válido...", vbExclamation
                  Exit Sub
                End If
                
                If UCase(Right(.FileName, 3)) <> "TXT" Then
                  MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                  Exit Sub
                End If
        
         txtArchivo.Text = .FileName
        
        End With

  Case "cargar"
    Call sbCargaTrama
  
End Select


End Sub


