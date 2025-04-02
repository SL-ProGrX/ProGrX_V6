VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_TramasINS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Envío / Recepción de Tramas"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab ssTab 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Tag             =   "u"
      Top             =   1200
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   11668
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   988
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
      TabPicture(0)   =   "frmSeguros_Tramas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtArchivo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtpVence"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpCuota"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "vGrid"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "tlbProceso"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "tlbX"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCambio"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNoExiste"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtExiste"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCasos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtMonto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Envío"
      TabPicture(1)   =   "frmSeguros_Tramas.frx":05FF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "tlbTrama"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cboRemesa"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lsw"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin MSComctlLib.ListView lsw 
         Height          =   4935
         Left            =   -73560
         TabIndex        =   20
         Top             =   840
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   8705
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
         NumItems        =   5
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
            Text            =   "Cobrado"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Mnt x Cobrar"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Diferencia"
            Object.Width           =   3246
         EndProperty
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
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   9495
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
         TabIndex        =   7
         ToolTipText     =   "Monto"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox txtCasos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   6
         Top             =   5520
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
         TabIndex        =   5
         Top             =   5520
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
         TabIndex        =   4
         Top             =   5520
         Width           =   975
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
         TabIndex        =   3
         Top             =   5520
         Width           =   975
      End
      Begin MSComctlLib.Toolbar tlbX 
         Height          =   660
         Left            =   10320
         TabIndex        =   1
         Top             =   120
         Width           =   405
         _ExtentX        =   714
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
               ImageIndex      =   9
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cargar"
               Object.ToolTipText     =   "Cargar información"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbProceso 
         Height          =   336
         Left            =   8400
         TabIndex        =   8
         Top             =   5400
         Width           =   2484
         _ExtentX        =   4392
         _ExtentY        =   582
         ButtonWidth     =   1852
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
               ImageIndex      =   10
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "Eliminar"
               Object.ToolTipText     =   "Elimina Trama Pendiente"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3735
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   6588
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
         MaxCols         =   495
         SpreadDesigner  =   "frmSeguros_Tramas.frx":0C00
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbTrama 
         Height          =   1320
         Left            =   -66120
         TabIndex        =   21
         Top             =   840
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   2328
         ButtonWidth     =   1826
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Archivo"
               Key             =   "Archivo"
               Object.ToolTipText     =   "Genera archivo de respuesta al INS"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cobros"
               Key             =   "Cobros"
               Object.ToolTipText     =   "Cobros realizados derivados de la trama"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cierra"
               Key             =   "Cierra"
               Object.ToolTipText     =   "Cierra Trama para Generar Informe"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Informe"
               Key             =   "Informe"
               ImageIndex      =   13
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCuota 
         Height          =   330
         Left            =   3120
         TabIndex        =   24
         Top             =   960
         Width           =   1335
         _Version        =   1441792
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   330
         Left            =   8880
         TabIndex        =   25
         Top             =   960
         Width           =   1335
         _Version        =   1441792
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   735
         Left            =   1320
         TabIndex        =   28
         Top             =   120
         Width           =   8895
         _Version        =   1441792
         _ExtentX        =   15690
         _ExtentY        =   1296
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
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
         Left            =   -74760
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -74760
         X2              =   -64200
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Remesa"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1575
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
         TabIndex        =   16
         Top             =   960
         Width           =   2655
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
         TabIndex        =   15
         Top             =   960
         Width           =   1815
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
         TabIndex        =   14
         Top             =   5520
         Width           =   855
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
         TabIndex        =   13
         Top             =   5280
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
         TabIndex        =   12
         Top             =   5280
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
         TabIndex        =   11
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Activa"
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
         TabIndex        =   10
         Top             =   5280
         Width           =   975
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
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
   End
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
      Top             =   0
      Visible         =   0   'False
      Width           =   2100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
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
            Picture         =   "frmSeguros_Tramas.frx":12DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":7B3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":E3A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":14C03
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1B465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1BC27
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1C2EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1CCDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1D69D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1E0BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1E893
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1F250
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Tramas.frx":1F947
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   330
      Left            =   5400
      TabIndex        =   26
      Top             =   360
      Width           =   5895
      _Version        =   1441792
      _ExtentX        =   10398
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
   Begin XtremeSuiteControls.FlatEdit txtTramaId 
      Height          =   315
      Left            =   9480
      TabIndex        =   27
      Top             =   720
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Trama ID..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   7440
      TabIndex        =   23
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aseguradora..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   3240
      TabIndex        =   22
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSeguros_TramasINS"
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



Private Function fxTramaCampo(pTrama As String, pCampo As Integer) As String
Dim vResultado As String, i As Integer, x As Integer, vChar As String

i = 0
vResultado = ""
x = 0
Do While i < pCampo
   x = x + 1
   vChar = Mid(pTrama, x, 1)
   If vChar = "," Or x = Len(pTrama) Then
      i = i + 1
      If x = Len(pTrama) Then vResultado = vResultado & vChar
   Else
      vResultado = vResultado & vChar
   End If
 
   If vChar = "," And i < pCampo Then
      vResultado = ""
   End If
 
Loop

fxTramaCampo = vResultado

End Function



Private Sub sbCargaTrama()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strCadena As String, curMonto As Currency
Dim fn, Casos(4) As Long, Campos() As String


Dim pTipoPoliza As String, pPoliza As String, pLinea As Long, pMoneda As String, pAseguradora As String
Dim pTipoId As String, pCedula As String, pNombre As String, pNumCta As Integer, pMonto As Currency
Dim pExiste As Integer, pTarjetaNum As String, pTarjetaVence As String, pComision As Currency
Dim pMedioPago As String, pFechaInicio As Date, pFechaCorte As Date, pTempo As String

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

pLinea = 0
strSQL = ""
pAseguradora = cboAseguradora.ItemData(cboAseguradora.ListIndex)

fn = FreeFile
Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
 Do While Not EOF(fn)
   Line Input #fn, strCadena
   Campos = Split(strCadena, ",")
    
' Script para Leer todas las columnas
'    For Columna = 0 To UBound(Campos)
'      x  = Campos(Columna)
'    Next Columna
    
   pLinea = pLinea + 1
   pNumCta = Campos(0)          'fxTramaCampo(strCadena, 1)
   pMoneda = Campos(1)          'fxTramaCampo(strCadena, 2)
   pPoliza = Campos(2)          'fxTramaCampo(strCadena, 3)
   pTarjetaNum = Campos(4)      'fxTramaCampo(strCadena, 5)
   pTarjetaVence = Campos(5)    'fxTramaCampo(strCadena, 6)
   pMedioPago = Campos(6)       'fxTramaCampo(strCadena, 7)
   pTempo = Campos(8)           'fxTramaCampo(strCadena, 9)
   If IsDate(pTempo) Then
       pFechaInicio = pTempo
   Else
       pFechaInicio = dtpCuota.Value
   End If
   pTempo = Campos(9)           'fxTramaCampo(strCadena, 9)
   If IsDate(pTempo) Then
       pFechaCorte = pTempo
   Else
       pFechaCorte = dtpVence.Value
   End If
   
   pMonto = Campos(10)          'fxTramaCampo(strCadena, 11)
   pComision = Campos(12)       'fxTramaCampo(strCadena, 13)
   pTipoId = Trim(Campos(14))         'fxTramaCampo(strCadena, 14)
   pCedula = Campos(13)         'fxTramaCampo(strCadena, 15)
   pNombre = UCase(Campos(15))  'fxTramaCampo(strCadena, 16)
   

   strSQL = strSQL & Space(10) & "insert seguros_Tramas(cod_aseguradora,Trama_Id, Cod_Linea,Num_Cuota, Num_Poliza, Monto, Comision_Neta" _
                     & ",Tipo_Id, Cedula, Nombre, Fecha_Cuota, Fecha_Vence, Moneda" _
                     & ",Medio_pago, Tarjeta_Numero, Tarjeta_Vence, Trama_Original, Registro_Fecha, Registro_Usuario)" _
                     & " Values('" & pAseguradora & "','" & txtTramaId.Text & "'," & pLinea & "," & pNumCta & ",'" & pPoliza & "'," & pMonto & "," & pComision _
                     & ",'" & pTipoId & "','" & pCedula & "','" & pNombre & "','" & Format(pFechaInicio, "yyyy/mm/dd") & "','" & Format(pFechaCorte, "yyyy/mm/dd") _
                     & "','" & pMoneda & "','" & pMedioPago & "','" & pTarjetaNum & "','" & pTarjetaVence & "','" & strCadena _
                     & "', getdate(),'" & glogon.Usuario & "')"
                     
   If Len(strSQL) > 20000 Then
      Call ConectionExecute(strSQL)
      strSQL = ""
   End If
    
 Loop
Close #fn
        
'Cadena Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

'Verifica y Devuelve Resultado
Call sbTrama_Verificada_Load(pAseguradora, txtTramaId.Text, 1, 1)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbTrama_Verificada_Load(pAseguradora As String, pTrama As String _
                                    , Optional pVerifica As Integer = 1, Optional pResultados As Integer = 1)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, Casos(4) As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

curMonto = 0

Casos(0) = 0 'Total
Casos(1) = 0 'Existe
Casos(2) = 0 'No Existe
Casos(3) = 0 'Cambios

strSQL = "exec spSeguros_Trama_Verifica '" & pAseguradora & "','" & pTrama & "','" & glogon.Usuario & "',1,1"
Call OpenRecordSet(rs, strSQL)

With vGrid

    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      .Col = 1
      .Text = rs!Cedula
      .Col = 2
      .Text = rs!Nombre
      .Col = 3
      .Text = CStr(rs!Num_Cuota)
      .Col = 4
      .Text = CStr(rs!Monto)
      .Col = 5
      .Text = rs!Num_Poliza
      .Col = 6
      .Text = rs!COD_PRODUCTO & ""
      .Col = 7
      
      Select Case rs!Existe
         Case 0 'No Existe
            Casos(2) = Casos(2) + 1
            .Value = vbUnchecked
         Case 1 'Existe
            Casos(1) = Casos(1) + 1
            curMonto = curMonto + rs!Monto
            .Value = vbChecked
         Case 2 'No se encuentra Actva
            Casos(3) = Casos(3) + 1
            .Value = -1
      End Select
      
      rs.MoveNext
    Loop
    rs.Close
        
End With
        
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
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtExiste.Text = 0
    txtNoExiste.Text = 0
    txtCambio.Text = 0
End Sub


Private Sub sbCargaTrama_OLD()
Dim strCadena As String, curMonto As Currency
Dim fn, Casos(4) As Long, pMonto As Currency
Dim strSQL As String, rs As New ADODB.Recordset

Dim pTipoPoliza As String, pPoliza As String

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
   pTipoPoliza = ""
   pPoliza = ""
   
   If Len(strCadena) >= 66 Then 'Largo de la Trama
            With vGrid
                
               .MaxRows = .MaxRows + 1
               .Row = .MaxRows
               
               .Col = 1 'Cedula
               .Text = fxRemueveCeroIzq(Mid(strCadena, 38, 10))
               
               .CellTag = strCadena
               
               
               .Col = 3 'No. Cuota
               .Text = fxRemueveCeroIzq(Mid(strCadena, 48, 4))
               
               pMonto = CCur(fxRemueveCeroIzq(Right(strCadena, 10))) / 100
               curMonto = curMonto + pMonto
             
               .Col = 4 'Monto
               .Text = pMonto
               
               .Col = 5 'No. Póliza
               .Text = Mid(strCadena, 16, 9)
               
                pPoliza = .Text
               
               .Col = 6 'Tipo de Seguro
               .Text = Mid(strCadena, 1, 11)
               
               Select Case .Text
                  Case "02640010026"
                    .Text = "02"
                  Case "02840010023"
                    .Text = "03"
                  Case "02940010022"
                    .Text = "04"
                  Case Else
                     MsgBox "No Está mapeada!", vbExclamation
                    .Text = "01"
               End Select
               
               pTipoPoliza = .Text
               
               
               
               
                strSQL = "select S.num_Poliza,isnull(P.nombre,'** Persona no registrada **') as 'Nombre',S.Cuota " _
                       & " from SEGUROS_REGISTRO S left Join Socios P on S.cedula = P.cedula" _
                       & " where right(num_Poliza, 9) = '" & pPoliza & "' and cod_Producto = '" & pTipoPoliza _
                       & "' and S.cod_aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
                
                Call OpenRecordSet(rs, strSQL)
                If Not rs.EOF And Not rs.BOF Then
                  .Col = 2 'Nombre
                  .Text = rs!Nombre
                  
                  .Col = 5
                  .Text = Trim(rs!Num_Poliza)
                
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
                
                'Totales
                txtMonto.Text = Format(curMonto, "Standard")
                txtCasos.Text = vGrid.MaxRows
                
                txtExiste.Text = Casos(1)
                txtNoExiste.Text = Casos(2)
                txtCambio.Text = Casos(3)
                DoEvents
            
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
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtExiste.Text = 0
    txtNoExiste.Text = 0
    txtCambio.Text = 0
End Sub



Private Sub cboAseguradora_Click()
Dim vFecha As Date

On Error GoTo vError

vGrid.MaxCols = 7
vGrid.MaxRows = 0

vFecha = fxFechaServidor

txtTramaId.Text = Format(vFecha, "yyyy.mm.dd")

ssTab.Tab = 0
dtpCuota.Value = vFecha
Call dtpCuota_Change

vError:

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

vModulo = 17

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 17

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

vPaso = True
    strSQL = "select cod_aseguradora as 'Idx', rtrim(nombre) as 'ItmX' from seguros_Aseguradoras"
    Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)
vPaso = False

Call cboAseguradora_Click

End Sub

Private Sub sbProcesar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vRemesa As Long

On Error GoTo vError

If vGrid.MaxRows = 0 Then
    MsgBox "No existen datos para procesar!"
    Exit Sub
End If

Me.MousePointer = vbHourglass


'spSeguros_Trama_Procesa(@Aseguradora varchar(10),@TramaId varchar(30), @FechaCuota datetime, @FechaVence datetime, @Usuario varchar(30))

strSQL = "exec spSeguros_Trama_Procesa '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtTramaId.Text & "','" _
        & Format(dtpCuota.Value, "yyyy/mm/dd") & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & " 23:59:59','" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)
 
vRemesa = rs!Remesa

rs.Close

Me.MousePointer = vbDefault

MsgBox "Carga de Seguros para el cobro realizado Satisfactoriamente -> Remesa: " & vRemesa, vbInformation

txtArchivo.Text = ""
vGrid.MaxRows = 0

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbProcesar_OLD()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCedula As String, pNombre As String, pCtaNum As Integer, pMonto As Currency, pPolizaNum As String, pPolizaSeg As String
Dim lng As Long, vProcesados As Long, vLinea As Integer
Dim pTramaId As String, pStrTrama As String, pTramaExec As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vProcesados = 0
strSQL = ""
pTramaExec = ""

pTramaId = Format(fxFechaServidor, "yyyy.mm.dd")

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
               
            strSQL = strSQL & Space(10) & "exec spSeguros_Trama_Procesa '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & pPolizaNum _
                   & "','" & pPolizaSeg & "'," & pMonto & "," & pCtaNum & ",'" _
                   & Format(dtpCuota.Value, "yyyy/mm/dd") & "','" & Format(dtpVence.Value, "yyyy/mm/dd") _
                   & "','" & pTramaId & "','" & glogon.Usuario & "'"
                   
            'Ejecuta Lote
            If Len(strSQL) > 20000 Then
               Call ConectionExecute(strSQL)
               strSQL = ""
            End If
        
        Else
         'Trama Inconsistente
            .Col = 1
            pCedula = Trim(.Text)
            pStrTrama = Trim(.CellTag)
            
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
               
            pTramaExec = pTramaExec & Space(10) & "exec spSeguros_Trama_Procesa_Erronea '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & pPolizaNum _
                   & "','" & pPolizaSeg & "'," & pMonto & "," & pCtaNum & ",'" _
                   & Format(dtpCuota.Value, "yyyy/mm/dd") & "','" & Format(dtpVence.Value, "yyyy/mm/dd") _
                   & "','" & pTramaId & "','" & glogon.Usuario & "','" & pStrTrama & "','" & pCedula & "'"
                   
            'Ejecuta Lote
            If Len(pTramaExec) > 30000 Then
               Call ConectionExecute(pTramaExec)
               pTramaExec = ""
            End If
                   
       
        End If '.Value = vbChecked
    
    Next lng
    
    'Ejecuta Ultimo Lote
    If Len(strSQL) > 0 Then
       Call ConectionExecute(strSQL)
       strSQL = ""
    End If

    'Ejecuta Lote
    If Len(pTramaExec) > 20000 Then
       Call ConectionExecute(pTramaExec)
       pTramaExec = ""
    End If

End With

Me.MousePointer = vbDefault

MsgBox "Carga de Seguros para el cobro...aplicadas Satisfactoriamente... Registros Procesados :" & vProcesados, vbInformation

txtArchivo.Text = ""
vGrid.MaxRows = 0

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbLswTramas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass
 

strSQL = "  select St.TRAMA_ID, COUNT(*) as 'Casos', isnull(SUM(Sp.MONTO_PAGO),0) as 'Monto', SUM(St.MONTO) as 'CxC' " _
       & "  from SEGUROS_REMESAS Sr" _
       & "      inner join SEGUROS_TRAMAS St  on Sr.TRAMA_ID = St.TRAMA_ID and St.COD_ASEGURADORA = Sr.COD_ASEGURADORA" _
       & "       left join SEGUROS_PAGOS Sp on Sr.COD_REMESA = Sp.COD_REMESA and Sr.TRAMA_ID = Sp.TRAMA_ID" _
       & "             and Sp.COD_ASEGURADORA = Sr.COD_ASEGURADORA" _
       & "             and Sp.NUM_POLIZA = St.NUM_POLIZA" _
       & "             and Sp.NUM_CUOTA = St.NUM_CUOTA" _
       & " Where Sr.cod_remesa = " & cboRemesa.ItemData(cboRemesa.ListIndex) _
       & " group by St.TRAMA_ID"

lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Trama_Id)
     itmX.SubItems(1) = Format(rs!Casos, "###,###,##0")
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.SubItems(3) = Format(rs!CxC, "Standard")
     itmX.SubItems(4) = Format(rs!CxC - rs!Monto, "Standard")
     itmX.Checked = True

 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset


If ssTab.Tab = 1 Then
    vPaso = True
    
    cboRemesa.Clear
    
    'and estado in('C','T')
    strSQL = "select Top 100 * from SEGUROS_REMESAS where Tipo = 'A' order by fecha_inicio desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboRemesa.AddItem (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!FECHA_CORTE, "dd/mm/yyyy"))
      cboRemesa.ItemData(cboRemesa.NewIndex) = rs!cod_remesa
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboRemesa.Text = (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!FECHA_CORTE, "dd/mm/yyyy"))
    End If
    rs.Close
    vPaso = False


   Call sbLswTramas
End If

End Sub


Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

On Error GoTo vError

Select Case Button.Key
  Case "aplicar"
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen casos a procesar...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
  
  Case "Eliminar"
  
    strSQL = "exec spSeguros_Trama_Elimina '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtTramaId.Text _
            & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
    
    txtArchivo.Text = ""
    Call cboAseguradora_Click
    
    Call Bitacora("Elimina", "Trama: " & txtTramaId.Text & " Aseguradora: " & cboAseguradora.ItemData(cboAseguradora.ListIndex))
    MsgBox "Trama Eliminada Satisfactoriamente!", vbInformation
    
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCreaTrama(vRemesa As Long, vTramaId As String)
Dim i As Long, vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String
Dim fnFile

Dim strSQL As String, rs As New ADODB.Recordset, lngMonto As Long

fnFile = FreeFile


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Seguros"

MkDir SIFGlobal.DirectorioDeResultados & "\Seguros\" & cboAseguradora.ItemData(cboAseguradora.ListIndex)

vRuta = SIFGlobal.DirectorioDeResultados & "\Seguros\" & cboAseguradora.ItemData(cboAseguradora.ListIndex)

vArchivo = "R_" & vTramaId & "_" & Format(vRemesa, "000") & ".txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError

Open vTempo For Output As #fnFile  ' Create file name.


'TODO: Agrupar por Formatos de Tramas
'spSeguros_Trama_Respuesta(@Aseguradora varchar(10), @Remesa int, @TramaId varchar(30), @Usuario varchar(30))

strSQL = "exec spSeguros_Trama_Respuesta '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'," & vRemesa & ",'" & vTramaId & "','" _
       & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vCadena = rs!Trama_Respuesta
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
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxTramaEstado(vRemesa As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select Estado from seguros_remesas where cod_remesa = " & vRemesa
Call OpenRecordSet(rs, strSQL)
If rs.BOF And rs.EOF Then
   strSQL = "C"
Else
   strSQL = rs!estado
End If
rs.Close

fxTramaEstado = strSQL


End Function



Private Sub tlbTrama_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String
Dim vEstado As String, vRemesa As Long

On Error GoTo vError

vRemesa = cboRemesa.ItemData(cboRemesa.ListIndex)
vEstado = fxTramaEstado(vRemesa)

Select Case Button.Key

  Case "Archivo" 'Genera Archivo de Respuesta al INS
        If vEstado = "A" Then
           MsgBox "Esta Remesa/Trama se encuentra abierta no puede generar archivo de respuesta!", vbExclamation
           Exit Sub
        End If
        
        For i = 1 To lsw.ListItems.Count
           If lsw.ListItems.Item(i).Checked Then
             Call sbCreaTrama(vRemesa, lsw.ListItems.Item(i).Text)
           End If
        Next i
 
 Case "Cobros"
            Me.MousePointer = vbHourglass

            strSQL = "exec spSeguros_CobrosActualiza_Lite '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
            Call ConectionExecute(strSQL)
            
            Me.MousePointer = vbDefault
            
            MsgBox "Cobranza de Cobros de los seguros, actualizada!", vbInformation

 
 Case "Cierra"
        If vEstado = "A" Then
        
        
            Me.MousePointer = vbHourglass
                'Actualiza Cobranza antes del Cierre de Remesa
                strSQL = "exec spSeguros_CobrosActualiza_Lite '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
                Call ConectionExecute(strSQL)
                
                'Cierre de Remesa y Actualiza Datos de Comisiones
                strSQL = "exec spSeguros_RemesaCierre " & vRemesa
                Call ConectionExecute(strSQL)
            Me.MousePointer = vbDefault
            
            MsgBox "Remesa/Trama Cerrada Satisfactoriamente!", vbInformation
        
        Else
           MsgBox "Esta Remesa/Trama NO se encuentra abierta posiblemente haya sido cerrada anteriormente!", vbExclamation
        End If
 
 Case "Informe"
       MsgBox "En Desarrollo!", vbInformation
        
        
 Case Else

End Select

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "buscar"
        
        txtArchivo.Text = ""
        
        With frmContenedor.CD
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



Private Sub txtTramaId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Trama_ID"
   gBusquedas.Orden = "Trama_ID"
   gBusquedas.Filtro = " and cod_aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) _
            & "' group by Trama_ID"
   gBusquedas.Consulta = "Select Trama_ID from SEGUROS_TRAMAS"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTramaId.Text = gBusquedas.Resultado
      Call sbTrama_Verificada_Load(cboAseguradora.ItemData(cboAseguradora.ListIndex), txtTramaId.Text, 1, 1)
   End If
End If
End Sub
