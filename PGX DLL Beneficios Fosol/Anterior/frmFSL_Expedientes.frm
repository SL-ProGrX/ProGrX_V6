VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Expedientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fosol: Expediente"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   13455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboDocumento 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox cboTipo 
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
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox cboCausa 
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
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtNumDocumento 
      Alignment       =   2  'Center
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
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   15
      ToolTipText     =   "Cédula Solicitante"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información Solicitante"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   13215
      Begin VB.TextBox txtPresentaContacto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   12975
      End
   End
   Begin VB.TextBox txtPresentaCedula 
      Alignment       =   2  'Center
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
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   "Cédula Solicitante"
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtPresentanombre 
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
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "Nombre Solicitante"
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtExpediente 
      Alignment       =   2  'Center
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
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   4
      ToolTipText     =   "Expediente"
      Top             =   840
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Créditos"
      TabPicture(0)   =   "frmFSL_Expedientes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vgCreditos"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle Causa"
      TabPicture(1)   =   "frmFSL_Expedientes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtEnfermedad"
      Tab(1).Control(1)=   "txtObservaciones"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Label9"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Requisitos"
      TabPicture(2)   =   "frmFSL_Expedientes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lsw"
      Tab(2).Control(1)=   "lblEstadoRequisitos"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Resolución"
      TabPicture(3)   =   "frmFSL_Expedientes.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtResolucionNotas"
      Tab(3).Control(1)=   "dtpFechaResolucion"
      Tab(3).Control(2)=   "dtpFechaCausa"
      Tab(3).Control(3)=   "lswComite"
      Tab(3).Control(4)=   "Label4(1)"
      Tab(3).Control(5)=   "Label4(2)"
      Tab(3).Control(6)=   "Label14"
      Tab(3).Control(7)=   "Label6"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Seguimiento"
      TabPicture(4)   =   "frmFSL_Expedientes.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "vGridHistoricoGestiones"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtResolucionNotas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   -73920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   33
         ToolTipText     =   "Observaciones"
         Top             =   1200
         Width           =   6255
      End
      Begin VB.TextBox txtEnfermedad 
         Alignment       =   2  'Center
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
         Left            =   -73440
         MaxLength       =   15
         TabIndex        =   24
         ToolTipText     =   "Cédula Solicitante"
         Top             =   540
         Width           =   3975
      End
      Begin VB.TextBox txtObservaciones 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -73440
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   1020
         Width           =   10335
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   3375
         Left            =   -72600
         TabIndex        =   17
         Top             =   1080
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin FPSpreadADO.fpSpread vGridHistoricoGestiones 
         Height          =   3780
         Left            =   360
         TabIndex        =   34
         Top             =   480
         Width           =   12495
         _Version        =   524288
         _ExtentX        =   22040
         _ExtentY        =   6668
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmFSL_Expedientes.frx":008C
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpFechaResolucion 
         Height          =   330
         Left            =   -72120
         TabIndex        =   35
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   151977987
         CurrentDate     =   41023
      End
      Begin MSComCtl2.DTPicker dtpFechaCausa 
         Height          =   330
         Left            =   -63840
         TabIndex        =   36
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   151977987
         CurrentDate     =   41023
      End
      Begin MSComctlLib.ListView lswComite 
         Height          =   3135
         Left            =   -66720
         TabIndex        =   37
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
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
         NumItems        =   0
      End
      Begin FPSpreadADO.fpSpread vgCreditos 
         Height          =   3855
         Left            =   -73800
         TabIndex        =   41
         Top             =   480
         Width           =   10455
         _Version        =   524288
         _ExtentX        =   18441
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
         MaxCols         =   7
         SpreadDesigner  =   "frmFSL_Expedientes.frx":073A
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de la resolución"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -73920
         TabIndex        =   43
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   40
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha en la que se establece la Causa"
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
         Left            =   -66840
         TabIndex        =   39
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Cómite"
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
         Left            =   -67440
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblEstadoRequisitos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71280
         TabIndex        =   28
         Top             =   540
         Width           =   5055
      End
      Begin VB.Label Label11 
         Caption         =   "Enfermedad"
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Observaciones"
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
         Left            =   -74880
         TabIndex        =   23
         Top             =   1020
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker dtpFechaExpediente 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   151977987
      CurrentDate     =   40640
   End
   Begin MSComctlLib.ImageList imgToolAux 
      Left            =   13440
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":0F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":7786
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":DFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":1484A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":1B0AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":1B1C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":21A28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   13440
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":2828A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":2EAEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":3534E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":3BBB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":42412
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":48C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":5DDE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1349
      BandCount       =   5
      _CBWidth        =   13455
      _CBHeight       =   765
      _Version        =   "6.7.9782"
      Child1          =   "txtCedula"
      MinHeight1      =   315
      Width1          =   2340
      NewRow1         =   0   'False
      MinHeight2      =   315
      Width2          =   5640
      NewRow2         =   0   'False
      Child3          =   "txtEstado"
      MinWidth3       =   165
      MinHeight3      =   315
      Width3          =   30
      NewRow3         =   0   'False
      Child4          =   "tlbMantenimiento"
      MinHeight4      =   330
      Width4          =   4095
      NewRow4         =   -1  'True
      MinHeight5      =   360
      NewRow5         =   0   'False
      Begin VB.TextBox txtCedula 
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
         Left            =   165
         MaxLength       =   15
         TabIndex        =   14
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   30
         Width           =   2145
      End
      Begin MSComctlLib.Toolbar tlbMantenimiento 
         Height          =   330
         Left            =   165
         TabIndex        =   13
         Top             =   390
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   582
         ButtonWidth     =   2117
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo Expediente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Key             =   "Guardar"
               Object.ToolTipText     =   "Guarda los datos"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Apelación"
               Key             =   "Apelacion"
               Object.ToolTipText     =   "Borrar Expediente"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               Object.ToolTipText     =   "imprimi Boleta del Expediente"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Gestión"
               Key             =   "Gestiones"
               ImageIndex      =   2
            EndProperty
         EndProperty
         Begin VB.Label Label5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   32
            Top             =   2160
            Width           =   5055
         End
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   30
         Width           =   5265
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8205
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   11
         ToolTipText     =   "Estado del Expediente"
         Top             =   30
         Width           =   5160
      End
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   12000
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":747A8
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":8991A
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":9EA8C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":B3BFE
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":C8D70
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":C964A
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":DE7BC
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F392E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F4208
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   12600
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F4AE2
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F4C00
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F4D26
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F4E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F4F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F5079
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F517A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F52B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F53C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F54EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expedientes.frx":F5613
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2160
      TabIndex        =   26
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   175439875
      CurrentDate     =   40640
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   8790
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Agencia / Oficina"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Total Deuda Resolución"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Total Deuda Presentaci{on"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMembresia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   31
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Cédula Solicitante"
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
      Left            =   120
      TabIndex        =   30
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image imgTags 
      Height          =   240
      Left            =   3840
      Picture         =   "frmFSL_Expedientes.frx":F59FB
      ToolTipText     =   "Etiquetas"
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha Establece la Causa"
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
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Casos de Liquidación"
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
      Left            =   8280
      TabIndex        =   21
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Causa"
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
      Left            =   8280
      TabIndex        =   20
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Documento Ref"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image imgNo 
      Height          =   255
      Left            =   360
      Picture         =   "frmFSL_Expedientes.frx":F5B05
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSi 
      Height          =   255
      Left            =   0
      Picture         =   "frmFSL_Expedientes.frx":F5C1F
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "Nombre Solicitante"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Expediente"
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha Expediente"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmFSL_Expedientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vFecha As String, vMembresiaMeses As Integer
Dim vRequisitosCompletos As Integer, vIdSolicitud As Long
Dim vTipoAplicacion As String, vEstadoRequisitos As String
Dim vPorcentaje As Double, vPorcentajeApl As Double
Dim vSaldo As Currency, vMORA_INT_COR As Currency, vMORA_INT_MOR As Currency, vCargos As Currency
Dim vMONTO_FOSOL As Currency, vPrimera_Deduccion As String, vMonto As Currency, vPaso As Boolean
Dim vDeuda_P As Currency, vEstado As String, vNuevo As Boolean, vCodExpediente As Currency

Private Sub cboTipo_Click()
  If vPaso Then Exit Sub
  Call sbCreditos
  Call sbCargaCausas(cboTipo.ItemData(cboTipo.ListIndex))
End Sub

Private Sub Form_Activate()
 vModulo = 22
End Sub

Private Sub sbCargaRequisitos()
On Error GoTo vError

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vItem As MSComctlLib.ListItem
Dim vKey As String
Dim vLvw As MSComctlLib.ListView

If txtExpediente.Text = "" Then Exit Sub

Me.lsw.ColumnHeaders.Clear
Me.lsw.ListItems.Clear
  
If cboCausa.Text = Empty Then Exit Sub
  
Set vLvw = Me.lsw
    vLvw.ColumnHeaders.Add , , "Requisito", 2500
    vLvw.ColumnHeaders.Add , , "Opcional", 1000, 2
    vLvw.ColumnHeaders.Add , , "Asignado", 1000, 2


'Trae los requisitos y checked los que estan asignados
strSQL = "Select distinct(RC.COD_REQUISITO), R.DESCRIPCION,RC.OPCIONAL,RC.ASIGNADO,ER.ESTADO as 'Presentado' " _
       & " from FSL_REQUISITOS_CAUSAS RC " _
       & " inner join FSL_REQUISITOS R on R.COD_REQUISITO = RC.COD_REQUISITO " _
       & " left join FSL_EXPEDIENTES_REQUISITOS ER on RC.COD_REQUISITO = ER.COD_REQUISITO" _
       & " and ER.COD_EXPEDIENTE = " & txtExpediente.Text & "" _
       & " Where RC.COD_CAUSA=" & cboCausa.ItemData(cboCausa.ListIndex) & " and RC.ASIGNADO <> 0  "

rs.Open strSQL, glogon.Conection, adOpenStatic

If rs.EOF Then
  MsgBox "No se tiene los requisitos configurados de esta Causa de Liquidación"
  rs.Close
  Exit Sub
End If

Do While Not rs.EOF
  vKey = Trim(rs.Fields("COD_REQUISITO")) & "(CA)"
  Set vItem = lsw.ListItems.Add(, vKey, Trim(rs.Fields!Descripcion))
              vItem.SubItems(1) = IIf(rs!Opcional = 1, "Si", "No")
              vItem.SubItems(2) = IIf(rs!Asignado = 1, "Si", "No")
  If rs!Presentado = 1 Then
     vItem.Checked = True
  End If
  rs.MoveNext
Loop

If fxRequisitosCompletos = True Then
   lblEstadoRequisitos.Caption = "Requisitos Básicos Completos"
   lblEstadoRequisitos.ForeColor = vbBlue
Else
   lblEstadoRequisitos.Caption = "Faltan requisitos Básicos por presentar"
   lblEstadoRequisitos.ForeColor = vbBlue
End If

rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
 Dim i As Integer
 Dim vOficina As String
  
  vModulo = 22
  vFecha = fxFechaServidor
  dtpFechaExpediente.Value = vFecha
  vPaso = True
  
  Call sbCargaCombos
  Call sbCargaTipos
  If cboTipo.ListCount > 0 Then
  Call sbCargaCausas(cboTipo.ItemData(cboTipo.ListIndex))
  End If
  ssTab.Tab = 0
  vgCreditos.MaxRows = 0
  
  vGridHistoricoGestiones.MaxRows = 0
  vOficina = GLOBALES.gOficinaTitular

End Sub

Private Sub imgTags_Click()
If txtExpediente.Text <> Empty Then
   GLOBALES.gTag = txtExpediente.Text
   GLOBALES.gTag2 = txtCedula.Text
   frmFSL_SeguimientoEtiquetas.Show vbModal
Else
   MsgBox "Debe de registrar el Expediente", vbInformation
End If
End Sub

Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vExiste As Boolean, i As Integer
Dim vCodRequisito As Integer, vOpcional As Integer
  
vExiste = True

vCodRequisito = DeCodificaPrimaryKey(Item.Key, 1, "(CA)")
strSQL = "select count(COD_REQUISITO) as 'Requisito' from FSL_EXPEDIENTES_REQUISITOS" _
       & " where COD_EXPEDIENTE = " & txtExpediente.Text & " and COD_REQUISITO = " & vCodRequisito & "  "
rs.Open strSQL, glogon.Conection, adOpenStatic
    
If rs!Requisito <= 0 Then
   vExiste = False
End If

rs.Close
  
If Item.Checked Then
  
  If Item.SubItems(1) = "Si" Then
      vOpcional = 1
  Else
      vOpcional = 0
  End If
    
  If vExiste = False Then
    strSQL = "INSERT FSL_EXPEDIENTES_REQUISITOS (COD_EXPEDIENTE,COD_REQUISITO,OPCIONAL,ESTADO,REGISTRO_FECHA,REGISTRO_USUARIO)" _
           & "Values(" & txtExpediente.Text & ",'" & vCodRequisito & "'," & vOpcional & ",1,'" & Format(vFecha, "yyyymmdd") & "','" & glogon.Usuario & "')"
    glogon.Conection.Execute strSQL
  Else
    strSQL = "UPDATE FSL_EXPEDIENTES_REQUISITOS SET OPCIONAL = " & vOpcional & ",ESTADO = 1,REGISTRO_FECHA = '" & Format(vFecha, "yyyymmdd") & "',REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
           & "WHERE COD_EXPEDIENTE=" & txtExpediente.Text & " and COD_REQUISITO='" & vCodRequisito & "'"
    glogon.Conection.Execute strSQL
  End If 'vExiste = False

Else
  
  If vExiste = True Then
    strSQL = "UPDATE FSL_EXPEDIENTES_REQUISITOS SET OPCIONAL = " & vOpcional & ",ESTADO = 0,REGISTRO_FECHA = '" & Format(vFecha, "yyyymmdd") & "',REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
           & "WHERE COD_EXPEDIENTE=" & txtExpediente.Text & " and COD_REQUISITO='" & vCodRequisito & "'"
    glogon.Conection.Execute strSQL
  End If 'vExiste = False

End If 'Item.Checked
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
   Select Case ssTab.Tab
      Case 2
        Call sbCargaRequisitos
      Case 3
        If txtExpediente.Text = Empty Then Exit Sub
        Call sbTraeResolucion(txtExpediente.Text)
      Case 4
        If txtExpediente.Text = Empty Then Exit Sub
        Call sbTraerHistoricoSeguimiento
   End Select
End Sub

Private Sub tlbMantenimiento_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Nuevo"
     Call sbLimpiar
     txtCedula.SetFocus
     
  Case "Guardar"
          
     If fxValidaExpediente Then
        
          If vNuevo Then
            Call sbGuardaExpediente
            Call sbGuardaDetalleExpediente
           
    '        strSQL = "exec spSifRegistraTags '" & txtCedula.Text & "','" & txtExpediente.Text & "', " _
    '               & " '','A04','" & glogon.Usuario & "','" & txtObservaciones.Text & "'," & txtExpediente.Text & ""
    '
    '        glogon.Conection.Execute strSQL
           
            Call Bitacora("Registra", "Registra Expediente FOSOL: Exp-" & txtExpediente.Text)
          Else
            Call sbModificaExpediente
            Call sbGuardaDetalleExpediente
           
    '        strSQL = "exec spSifRegistraTags '" & txtCedula.Text & "','" & txtExpediente.Text & "', " _
    '               & " '','A04','" & glogon.Usuario & "','" & txtObservaciones.Text & "'," & txtExpediente.Text & ""
    '
    '        glogon.Conection.Execute strSQL
            
            Call Bitacora("Modifica", "Modifica Expediente FOSOL: Exp-" & txtExpediente.Text)
          End If 'vNuevo = true
         
         If vRequisitosCompletos = 0 Then
           MsgBox "Información Guardada Satisfactoriamente" & vbCrLf & "Faltan requisitos para la aprobación", vbExclamation
         Else
           MsgBox "Información Guardada Satisfactoriamente", vbInformation
         End If
         
         txtExpediente.Text = vCodExpediente
     
     End If 'fxValidaExpediente = true
     
     Call sbTraeExpediente
  
  Case "Apelacion"
     If vEstado = "REC" Or vEstado = "APL" Then
        GLOBALES.gTag = txtCedula.Text
        GLOBALES.gTag2 = txtNombre.Text
        GLOBALES.gTag3 = txtExpediente.Text
        frmFSL_Apelacion.Show vbModal
     End If
  
  Case "Imprimir"
     If txtExpediente.Text = Empty Then Exit Sub
     Call BoletaFOSOL(txtExpediente.Text)
     
  Case "Gestiones"
     If txtCedula.Text = Empty Or txtExpediente.Text = Empty Then Exit Sub
     GLOBALES.gTag = txtCedula.Text
     GLOBALES.gTag2 = txtNombre.Text
     GLOBALES.gTag3 = txtExpediente.Text
     frmFSL_SeguimientoEtiquetas.Show vbModal
   
  Case "Datos"
     If txtCedula.Text = Empty Then Exit Sub
'     Call sbSIFForms("frmAF_Principal")
     
End Select
  
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCedTemp As String

On Error GoTo vError

'Busca primer en el Maestro de Socios, de lo contrario revisa si es una operacion
' y regresa la cedula de la operacion
vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 Call sbLimpiar
 strSQL = "select coalesce(count(*),0) as Existe from socios where cedula = '" & txtCedula.Text & "'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then
   rs.Close
   strSQL = "select cedula from reg_creditos where id_solicitud = '" & txtCedula.Text & "'"
   rs.Open strSQL, glogon.Conection, adOpenStatic
   If Not rs.EOF And Not rs.BOF Then
      vCedTemp = Trim(rs!Cedula)
   End If
 End If
 rs.Close

If vCedTemp = "" Then
  Call sbConsulta(txtCedula.Text)
  Call sbTraeExpediente
Else
  Call sbConsulta(vCedTemp)
End If

'Trae los datos Aprobados por el cómite
Call sbTraeDatosAprobacion
End If

If KeyCode = vbKeyF4 Then Call sbBusqueda

txtExpediente.Enabled = False



Exit Sub
vError:
    MsgBox Err.Description, vbCritical
    Resume
End Sub

Private Sub sbBusqueda()
On Error GoTo vError
  gBusquedas.Convertir = "N"
  gBusquedas.Consulta = "Select cedula,nombre from SOCIOS"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  frmBusquedas.Show vbModal
  txtCedula = Trim(gBusquedas.Resultado)
  gBusquedas.Consulta = ""
  gBusquedas.Columna = ""
  gBusquedas.Orden = ""
  gBusquedas.Resultado = ""
  If Trim(txtCedula) <> "" Then
      Call sbLimpiar
      Call sbConsulta(txtCedula)
  End If
   
  Call sbTraeExpediente
  
Exit Sub

vError:
  MsgBox Err.Description, vbCritical
Resume
End Sub
Private Sub txtEnfermedad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservaciones.SetFocus
End Sub

Private Sub txtExpediente_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtCedula.Text = Empty
  Call sbTraeExpediente
  cboTipo.SetFocus
End If

If KeyCode = vbKeyF4 Then
    txtCedula.Text = Empty
    gBusquedas.Columna = "COD_EXPEDIENTE"
    gBusquedas.Orden = "COD_EXPEDIENTE"
    gBusquedas.Filtro = ""
    gBusquedas.Consulta = "select COD_EXPEDIENTE,CEDULA from FSL_EXPEDIENTES"
    frmBusquedas.Show vbModal
    txtExpediente = gBusquedas.Resultado
    txtCedula = gBusquedas.Resultado2
    cboTipo.SetFocus
End If

End Sub

Private Sub sbTraeExpediente()
On Error GoTo vError

If txtCedula.Text = Empty And txtExpediente.Text = Empty Then Exit Sub

strSQL = "Select E.COD_EXPEDIENTE, E.COD_PLAN, T.DESCRIPCION as 'DES_TIPO', E.COD_CAUSA, C.DESCRIPCION AS 'DES_CAUSA', E.CEDULA,S.NOMBRE, E.DOCUMENTO_REF" _
       & ", E.NUMERO_DOC_REF,E.PRESENTA_CEDULA, E.PRESENTA_NOMBRE, E.PRESENTA_CONTACTO, E.MEMBRESIA_MESES, E.MEMBRESIA_PORCENTAJE" _
       & ", E.REQUISTOS_COMPLETOS, E.OBSERVACIONES, E.DETALLE_ENFERMEDAD, E.ENFERMEDAD_USUARIO,E.REGISTRA_OFICINA" _
       & ", E.REGISTRO_FECHA, E.REGISTRO_USUARIO, E.MODIFICA_USUARIO, E.MODIFICA_FECHA, E.ESTADO" _
       & ", E.RESOLUCION_ESTADO, E.RESOLUCION_NOTAS, E.RESOLUCION_FECHA, E.RESOLUCION_USUARIO, E.TOTAL_FOSOL" _
       & ", E.TOTAL_SOBRANTE from FSL_EXPEDIENTES E" _
       & "  inner join FSL_PLANES_APLICACION T on E.COD_PLAN = T.COD_PLAN" _
       & "  inner join FSL_CAUSAS C ON E.COD_CAUSA = C.COD_CAUSA" _
       & "  inner join SOCIOS S on E.CEDULA = S.CEDULA"

If txtCedula.Text = Empty Then
  strSQL = strSQL & " where E.COD_EXPEDIENTE = " & txtExpediente.Text & " "
Else
  strSQL = strSQL & " where E.CEDULA = '" & txtCedula.Text & "' "
End If

rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  txtExpediente.Text = rs!COD_EXPEDIENTE
  cboDocumento.Text = rs!DOCUMENTO_REF
  txtNumDocumento.Text = rs!NUMERO_DOC_REF
  dtpFechaExpediente.Value = rs!Registro_Fecha
  dtpFechaExpediente.Enabled = False
  txtPresentaCedula.Text = rs!Presenta_Cedula
  txtPresentaNombre.Text = rs!Presenta_Nombre
  txtPresentaContacto.Text = rs!PRESENTA_CONTACTO
  txtObservaciones.Text = rs!observaciones
  txtEnfermedad.Text = IIf(IsNull(rs!DETALLE_ENFERMEDAD), "No se especifica", rs!DETALLE_ENFERMEDAD)
  cboTipo.Text = rs!DES_TIPO
     
  
  Select Case rs!Estado
    Case "REC"
       txtEstado.Text = "Rechazada"

    Case "APR"
       txtEstado.Text = "Aprobada"

    Case "PEN"
       txtEstado.Text = "Pendiente"
       
    Case "APL"
       txtEstado.Text = "Apelación"
       
  End Select
  vEstado = rs!Estado
  
  StatusBarX.Panels(1).Text = glogon.Usuario
  StatusBarX.Panels(2).Text = GLOBALES.gOficinaTitular
 
  rs.MoveNext
  vPaso = False

End If

rs.Close
Call sbCargaCausas(cboTipo.ItemData(cboTipo.ListIndex))

Call sbConsulta(txtCedula.Text)

Exit Sub
vError:
  MsgBox Err.Description, vbCritical
  Resume
End Sub

Private Sub sbCargaCombos()
    cboDocumento.Clear
    cboDocumento.AddItem "Accion de personal"
    cboDocumento.AddItem "Acta defunción"
    cboDocumento.Text = "Accion de personal"
End Sub

Private Sub sbCargaTipos()
On Error GoTo vError

strSQL = "select COD_PLAN as idx,DESCRIPCION as itmx" _
      & " From FSL_PLANES_APLICACION Where ACTIVO = 1"

Call sbLlenaCbo(cboTipo, strSQL, False, True)
vPaso = False

Exit Sub
vError:
  MsgBox Err.Description
End Sub

Private Sub sbCargaCausas(Optional Tipo As Integer = Empty)
On Error GoTo vError
  
   
If vPaso Then Exit Sub

strSQL = "Select COD_CAUSA as idx,DESCRIPCION as itmx from FSL_CAUSAS " _
       & " where COD_PLAN = " & Tipo & ""

Call sbLlenaCbo(cboCausa, strSQL, False, True)
      
Exit Sub
vError:
   MsgBox Err.Description

End Sub

Private Sub sbGuardaExpediente()

strSQL = "INSERT FSL_EXPEDIENTES (COD_EXPEDIENTE, COD_PLAN, COD_CAUSA, CEDULA, DOCUMENTO_REF,NUMERO_DOC_REF, PRESENTA_CEDULA, PRESENTA_NOMBRE" _
       & ", PRESENTA_CONTACTO, MEMBRESIA_MESES, MEMBRESIA_PORCENTAJE, REQUISTOS_COMPLETOS, OBSERVACIONES, LIQUIDACION_ESTADO, LIQUIDACION_NUMERO" _
       & ", REGISTRO_FECHA, REGISTRO_USUARIO, ESTADO, DETALLE_ENFERMEDAD,REGISTRA_OFICINA,APLICADO)" _
       & "  Values (" & vCodExpediente & "," & cboTipo.ItemData(cboTipo.ListIndex) & "," & cboCausa.ItemData(cboCausa.ListIndex) & ",'" & txtCedula & "'" _
       & ", '" & cboDocumento.Text & "','" & txtNumDocumento.Text & "','" & txtPresentaCedula.Text & "','" & txtPresentaNombre.Text & "'" _
       & ", '" & txtPresentaContacto & "'," & vMembresiaMeses & ",0,'" & vRequisitosCompletos & "','" & txtObservaciones.Text & "','',''" _
       & ", getdate(),'" & glogon.Usuario & "','PEN','" & txtEnfermedad.Text & "','" & GLOBALES.gOficinaTitular & "','N')"
       
glogon.Conection.Execute strSQL

End Sub

Private Sub sbModificaExpediente()

strSQL = "UPDATE FSL_EXPEDIENTES SET DOCUMENTO_REF ='" & cboDocumento.Text & "' ,NUMERO_DOC_REF ='" & txtNumDocumento.Text & "',REGISTRO_FECHA = '" & Format(dtpFechaExpediente, "yyyymmdd") & "'" _
       & ", COD_PLAN ='" & cboTipo.ItemData(cboTipo.ListIndex) & "' ,COD_CAUSA ='" & cboCausa.ItemData(cboCausa.ListIndex) & "' " _
       & ", REQUISTOS_COMPLETOS ='" & vRequisitosCompletos & "',PRESENTA_CEDULA = '" & txtPresentaCedula.Text & "',PRESENTA_NOMBRE ='" & txtPresentaNombre.Text & "'" _
       & ", PRESENTA_CONTACTO = '" & txtPresentaContacto & "'" _
       & ", OBSERVACIONES ='" & txtObservaciones.Text & "' ,MODIFICA_USUARIO ='" & glogon.Usuario & "' " _
       & ", MODIFICA_FECHA = getdate(),DETALLE_ENFERMEDAD='" & UCase(txtEnfermedad.Text) & "'" _
       & " WHERE  CEDULA = '" & Trim(txtCedula.Text) & "' and COD_EXPEDIENTE = " & vCodExpediente & ""
       
glogon.Conection.Execute strSQL

End Sub

'Guarda el detalle de la aplicacion
Private Sub sbGuardaDetalleExpediente()
Dim i As Integer, vExiste As Boolean

vPorcentajeApl = 0
vMonto = 0
vSaldo = 0
vPorcentaje = 0
vMONTO_FOSOL = 0
vExiste = True

With vgCreditos

For i = 1 To .MaxRows
    vExiste = True
    .Row = i
    .Col = 1
    If .Text = "" Then Exit Sub
    vIdSolicitud = .Text
    .Col = 3 'primera deducción
    vPrimera_Deduccion = Format(.Text, "yyyymmdd")
    .Col = 4 'Monto
    vMonto = Format(.Text, "standard")
    .Col = 5 'Total Deuda
    vDeuda_P = Format(.Text, "standard")
    .Col = 7 'Porcentaje
    vPorcentaje = IIf(Format(.Text, "Standard") = Empty, 0, Format(.Text, "Standard"))
        
    strSQL = "Select COD_EXPEDIENTE,ID_SOLICITUD from FSL_EXPEDIENTES_DETALLE" _
           & " where COD_EXPEDIENTE = " & vCodExpediente & " and ID_SOLICITUD = " & vIdSolicitud & ""
    rs.Open strSQL, glogon.Conection, adOpenStatic
         
    If rs.EOF Then
      vExiste = False
    End If
    
    rs.Close
    
    If vExiste Then
       strSQL = "Update FSL_EXPEDIENTES_DETALLE set PORCENTAJE = " & vPorcentaje & "" _
              & ", ACTUALIZACION_FECHA = '" & Format(vFecha, "yyyymmdd") & "',MODIFICA_USUARIO = '" & glogon.Usuario & "'" _
              & ", MODIFICA_FECHA= '" & Format(vFecha, "yyyymmdd") & "', TIPO_APLICACION_FOSOL = '" & vTipoAplicacion & "'" _
              & ", PRIMERA_DEDUCCION='" & vPrimera_Deduccion & "' " _
              & " where COD_EXPEDIENTE = " & vCodExpediente & " and ID_SOLICITUD = " & vIdSolicitud & ""
      glogon.Conection.Execute strSQL
    Else
      strSQL = "INSERT FSL_EXPEDIENTES_DETALLE (COD_EXPEDIENTE,ID_SOLICITUD,MONTO_FORMALIZADO,PORCENTAJE, PRIMERA_DEDUCCION " _
             & ", ACTUALIZACION_FECHA,MODIFICA_USUARIO,MODIFICA_FECHA,TIPO_APLICACION_FOSOL,TOTAL_DEUDA_P)" _
             & " Values (" & vCodExpediente & "," & vIdSolicitud & "," & vMonto & "" _
             & ", " & vPorcentaje & ",'" & vPrimera_Deduccion & "','" & Format(vFecha, "yyyymmdd") & "'" _
             & ", '" & glogon.Usuario & "','" & Format(vFecha, "yyyymmdd") & "','" & vTipoAplicacion & "'," & vDeuda_P & ")"
      glogon.Conection.Execute strSQL
    End If 'Existe= true
Next i

End With

End Sub

Private Sub sbGuardaRequisitos(vExpediente As Integer)
Dim vExiste As Boolean, vCodRequisito As String
Dim i As Integer, vOpcional As Integer


With lsw
For i = 1 To .ListItems.Count
  
  vExiste = True
  
  vCodRequisito = DeCodificaPrimaryKey(.ListItems.Item(i).Key, 1, "(CA)")
  strSQL = "select count(COD_REQUISITO) as 'Requisito' from FSL_EXPEDIENTES_REQUISITOS" _
         & " where COD_EXPEDIENTE = " & vExpediente & " and COD_REQUISITO = " & vCodRequisito & "  "
  rs.Open strSQL, glogon.Conection, adOpenStatic
      
  If rs!Requisito <= 0 Then
     vExiste = False
  End If
  rs.Close
  
  If .ListItems.Item(i).Checked Then
    If .ListItems(i).SubItems(1) = "Si" Then
        vOpcional = 1
    Else
        vOpcional = 0
    End If
      
    If vExiste = False Then
      strSQL = "INSERT FSL_EXPEDIENTES_REQUISITOS (COD_EXPEDIENTE,COD_REQUISITO,OPCIONAL,ESTADO,REGISTRO_FECHA,REGISTRO_USUARIO)" _
             & "Values(" & vExpediente & ",'" & vCodRequisito & "'," & vOpcional & ",1,'" & Format(vFecha, "yyyymmdd") & "','" & glogon.Usuario & "')"
      glogon.Conection.Execute strSQL
    Else
      strSQL = "UPDATE FSL_EXPEDIENTES_REQUISITOS SET OPCIONAL = " & vOpcional & ",ESTADO = 1,REGISTRO_FECHA = '" & Format(vFecha, "yyyymmdd") & "',REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
             & "WHERE COD_EXPEDIENTE=" & vExpediente & " and COD_REQUISITO='" & vCodRequisito & "'"
      glogon.Conection.Execute strSQL
    End If 'vExiste = False
  Else '.ListItems.Item(i).Checked
    If vExiste = True Then
      strSQL = "UPDATE FSL_EXPEDIENTES_REQUISITOS SET OPCIONAL = " & vOpcional & ",ESTADO = 0,REGISTRO_FECHA = '" & Format(vFecha, "yyyymmdd") & "',REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
             & "WHERE COD_EXPEDIENTE=" & vExpediente & " and COD_REQUISITO='" & vCodRequisito & "'"
      glogon.Conection.Execute strSQL
    End If 'vExiste = False

  End If ' no(.ListItems.Item(1).Checked)
  
Next i
End With
End Sub

'Verdadero si Existe el expediente
Private Function fxValidaExpediente() As Boolean
On Error GoTo vError
    
fxValidaExpediente = True

If txtCedula.Text = Empty Or txtNombre.Text = Empty Then
  MsgBox "Faltan los datos del Asociado"
  fxValidaExpediente = False
  Exit Function
Else
  strSQL = "Select COD_EXPEDIENTE as 'Expediente' from FSL_EXPEDIENTES" _
         & " where Cedula = '" & txtCedula.Text & "'" _
         & " group by COD_EXPEDIENTE"
  rs.Open strSQL, glogon.Conection, adOpenStatic
     
  If Not (rs.EOF) Then
    vCodExpediente = rs!Expediente
    rs.Close
    vNuevo = False
  Else
    rs.Close
    
    strSQL = "Select isnull(max(COD_EXPEDIENTE),0) as 'Expediente' from FSL_EXPEDIENTES"
    rs.Open strSQL, glogon.Conection, adOpenStatic
  
    vCodExpediente = (rs!Expediente) + 1

    rs.Close
    
  End If
End If

If txtNumDocumento = Empty Then
  MsgBox "Falta el número del documento que se presenta", vbCritical
  fxValidaExpediente = False
  Exit Function
End If

If fxRequisitosCompletos = True Then
   vRequisitosCompletos = 1
   lblEstadoRequisitos.Caption = "Requisitos Básicos Completos"
   lblEstadoRequisitos.ForeColor = vbBlue
Else
   vRequisitosCompletos = 0
   lblEstadoRequisitos.Caption = "Faltan requisitos Básicos por presentar"
   lblEstadoRequisitos.ForeColor = vbBlue
End If

txtPresentaCedula.Text = IIf(txtPresentaCedula.Text = Empty, txtCedula.Text, txtPresentaCedula.Text)
txtPresentaNombre.Text = IIf(txtPresentaNombre.Text = Empty, txtNombre.Text, txtPresentaNombre.Text)
txtPresentaContacto.Text = IIf(txtPresentaContacto.Text = Empty, "Sin datos del Beneficiario", txtPresentaContacto.Text)
txtEnfermedad.Text = IIf(txtEnfermedad.Text = Empty, "Sin definicion", txtEnfermedad.Text)
txtObservaciones.Text = IIf(txtObservaciones.Text = Empty, "Sin Observaciones registradas", txtObservaciones.Text)

Exit Function
vError:
   MsgBox Err.Description

End Function

Private Sub sbLimpiar()
Dim Control As Control, i As Integer
Dim vFecha As String


vFecha = fxFechaServidor

txtNombre.Text = Empty
txtExpediente.Text = Empty
txtNumDocumento.Text = Empty
txtPresentaCedula.Text = Empty
txtPresentaNombre.Text = Empty
txtPresentaContacto.Text = Empty
dtpFechaCausa.Value = vFecha
dtpFechaExpediente.Value = vFecha

Call sbCargaCombos
Call sbCargaTipos
Call sbCargaCausas(cboTipo.ItemData(cboTipo.ListIndex))

txtExpediente.Enabled = False

With vgCreditos
 .MaxRows = 1
 For i = 1 To .MaxCols
   .Col = i
   .Text = ""
 Next i
End With

vNuevo = True

End Sub

Private Sub sbConsulta(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset
Dim vEstadoSocio As String
     
vFianzas = False
    
strSQL = "select cedula as CedulaX,nombre,fechaingreso,estadoactual" _
       & " from socios Where cedula = '" & Trim(pCedula) & "'"

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
 
If Not rs.EOF And Not rs.BOF Then
   
   txtCedula.Text = Trim(rs!Cedulax & "")
   txtNombre.Text = rs!Nombre & ""
     
   vFechaIng = IIf(IsNull(rs!FechaIngreso), fxFechaServidor, rs!FechaIngreso)
   lblMembresia.ForeColor = vbBlue
   lblMembresia.FontBold = False
    
   If rs!EstadoActual = "S" Then
      lblMembresia.Caption = "Membresía: " & fxMembresia(vFechaIng)
      lblMembresia.ToolTipText = "[Ing.:" & Format(vFechaIng, "dd/mm/yyyy") & "]"
              
      'Consulta si tiene renuncia en tramite
      strSQL = "select count(*) as Existe from afi_cr_renuncias where cedula = '" & pCedula & "' and estado = 'T'"
      rsTmp.Open strSQL, glogon.Conection, adOpenStatic
      If rsTmp!existe > 0 Then
         lblMembresia.Caption = " ** Renuncia en Transito ** " & lblMembresia.Caption
         lblMembresia.ForeColor = vbRed
         lblMembresia.FontBold = True
      End If
     rsTmp.Close
    Else
     lblMembresia.Caption = "Membresía: NADA"
    End If
   
   vMembresiaMeses = CInt(DateDiff("M", vFechaIng, fxFechaServidor))
   rs.Close
       
   ssTab.Tab = 0
 
 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
   

End Sub

Private Sub sbCreditos(Optional pSheet As Integer = 1)
Dim rs As New ADODB.Recordset, strSQL As String
Dim curCuota As Currency, curMonto As Currency
Dim curSaldo As Currency, vMora As Boolean
Dim i As Integer, vDisponible As Currency
Dim vSobrante As Currency, vBase As Currency, vTSobrante As Currency
Dim vPlan As Integer

On Error GoTo vError

curCuota = 0
curMonto = 0
curSaldo = 0
vMora = False
vPlan = cboTipo.ItemData(cboTipo.ListIndex)

Me.MousePointer = vbHourglass
vMora = False

With vgCreditos
 .MaxRows = 1
 For i = 1 To .MaxCols
   .Col = i
   .Text = ""
 Next i
 
 strSQL = "exec spFSL_Creditos '" & txtCedula.Text & "'," & vPlan & " "
 
 rs.CursorLocation = adUseServer
 rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
.Row = .MaxRows
vTipoAplicacion = rs!TipoAplicacion

If rs!TipoAplicacion = "M" Then
   vBase = rs!Monto
Else
   vBase = rs!TDeudaActual
End If

vDisponible = vBase * rs!PorcentajeAplicacion / 100
vSobrante = rs!TDeudaActual - vDisponible
vTSobrante = vTSobrante + vSobrante

For i = 1 To .MaxCols
  .Col = i
  Select Case i
    Case 1 'Operacion
       .Text = CStr(rs!id_Solicitud)
    
    Case 2 'Garantia
        .Text = rs!Garantia
        .TextTip = TextTipFixed
        .TextTipDelay = 1000
        .CellNote = rs!GarantiaDetalle
       
    Case 3 'Primer Deduccion
        .Text = Format(rs!Prideduc, "####-##")
    
    Case 4 'Monto
        .Text = Format(rs!Monto, "Standard")
        
    Case 5  'Total Deuda en el Momento de la Presentacion
      .Text = Format(rs!TDeudaActual, "Standard")
      
    Case 6 'Total de Deuda al momento de la resolución
      .Text = Format(0, "Standard")
    
    Case 7
      .Text = Format(rs!PorcentajeAplicacion, "Standard")
      
    Case 8 'Aplica % s/Monto
      .Value = IIf((rs!TipoAplicacion = "M"), 1, 0)
      
    Case 9 'Aplica % s/Total Deuda (Saldo)
      .Value = IIf((rs!TipoAplicacion <> "M"), 1, 0)
      
    Case 10  'Disponible
      .Text = Format(vDisponible, "Standard")
      
    Case 11  'Sobrante
      .Text = Format(vSobrante, "Standard")
      
  End Select
Next i

.MaxRows = .MaxRows + 1
rs.MoveNext

Loop

rs.Close
End With
Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub txtExpediente_KeyPress(KeyAscii As Integer)
  If (IsNumeric(Chr(KeyAscii)) <> True) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
     KeyAscii = 0
  End If
End Sub

Private Sub txtNumDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaCedula.SetFocus
End Sub

Private Sub txtPresentaCedula_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaNombre.SetFocus
End Sub

Private Sub txtPresentaNombre_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaContacto.SetFocus
End Sub

Public Function fxMembresia(vFecha As Date) As String
Dim iDias As Integer, iMes As Integer, iAnio As Integer
Dim vResultado As String

On Error GoTo vError

iDias = 0
iMes = 0
iAnio = 0
vResultado = ""

iDias = DateDiff("d", vFecha, fxFechaServidor)

Do While iDias > 365
  iAnio = iAnio + 1
  iDias = iDias - 365
Loop

Do While iDias > 30
  iMes = iMes + 1
  iDias = iDias - 30
Loop

If iAnio > 0 Then vResultado = vResultado & iAnio & " año(s)"
If iMes > 0 Then
  If Len(vResultado) > 0 Then vResultado = vResultado & ", "
  vResultado = vResultado & iMes & " mes(es)"
End If
  
If iDias > 0 Then
  If Len(vResultado) > 0 Then vResultado = vResultado & " con "
  vResultado = vResultado & iDias & " dia(s) "
End If
  
fxMembresia = vResultado

Exit Function

vError:
 fxMembresia = "Membresía no válida"

End Function

Private Function fxRequisitosCompletos() As Boolean
Dim i As Integer

fxRequisitosCompletos = True
With lsw
For i = 1 To .ListItems.Count
  
  If lsw.ListItems(i).SubItems(1) = "No" And lsw.ListItems(i).SubItems(2) = "Si" Then
    If Not (lsw.ListItems(i).Checked) Then
        fxRequisitosCompletos = False
    End If
  End If
  
  If lsw.ListItems(i).SubItems(1) = "Si" And lsw.ListItems(i).SubItems(2) = "Si" Then
    If Not (lsw.ListItems(i).Checked) Then
        fxRequisitosCompletos = True
    End If
  End If
  
Next i
End With

End Function

Private Sub sbTraeDatosAprobacion()
Dim i As Integer, vDisponible As Currency, vSobrante As Currency, vBase As Currency, vTSobrante As Currency

vTSobrante = 0

With vgCreditos
    
    If txtExpediente.Text = Empty Then Exit Sub
       strSQL = "Select ED.COD_EXPEDIENTE,G.descripcion as Garantia, ED.PRIMERA_DEDUCCION,ED.ID_SOLICITUD " _
           & ", ED.TOTAL_DEUDA_P, ED.MONTO_LIQUIDACION, ED.PORCENTAJE, ED.MONTO_FOSOL" _
           & ", ED.TIPO_APLICACION_FOSOL, ED.MONTO_FORMALIZADO, E.ESTADO,Reg.SALDO as 'SaldoActual'" _
           & ", isnull(Vm.IntC + Vm.IntM + Vm.Cargos + Vm.Poliza,0) + Reg.Saldo as 'TDeudaActual' " _
           & " from FSL_EXPEDIENTES E inner join FSL_EXPEDIENTES_DETALLE ED on ED.COD_EXPEDIENTE = E.COD_EXPEDIENTE" _
           & "  inner join reg_creditos Reg on Ed.id_Solicitud = Reg.Id_Solicitud" _
           & "  inner join Crd_Garantia_Tipos G on Reg.Garantia = G.Garantia " _
           & "  left join vista_morosidad Vm on Reg.id_solicitud = Vm.id_Solicitud" _
           & " where E.COD_EXPEDIENTE = '" & txtExpediente.Text & "' "
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    Do While Not rs.EOF
      .Row = .MaxRows
      
      If rs!Tipo_Aplicacion_Fosol = "M" Then
         vBase = rs!MONTO_FORMALIZADO
      Else
         vBase = IIf(IsNull(rs!TOTAL_DEUDA_P), 0, rs!TOTAL_DEUDA_P)
      End If
      
      vDisponible = vBase * rs!Porcentaje / 100
      vSobrante = vDisponible - rs!TDeudaActual
      vTSobrante = vTSobrante + vSobrante
      
      .Col = 1 'Operacion
      .Text = CStr(rs!id_Solicitud)
            
      .Col = 2
      .Text = rs!Garantia
      
      .Col = 3 'Primer Deduccion
      .Text = Format(Format(rs!PRIMERA_DEDUCCION, "yyyymm"), "####-##")
      
      .Col = 4 'Monto
      .Text = Format(rs!MONTO_FORMALIZADO, "Standard")
      
      .Col = 5 'Total Deuda en el Momento de la Presentacion
      .Text = Format(rs!TDeudaActual, "Standard")
      
      .Col = 6 'Total de Deuda al momento de la Resolución
      .Text = Format(rs!TDeudaActual, "Standard")
    
      .Col = 7
      .Text = Format(rs!Porcentaje, "Standard")
      
      .Col = 8 'Aplica % s/Monto
      .Value = IIf((rs!Tipo_Aplicacion_Fosol = "M"), 1, 0)
      
      .Col = 9 'Aplica % s/Total Deuda (Saldo)
      .Value = IIf((rs!Tipo_Aplicacion_Fosol <> "M"), 1, 0)
      
      .Col = 10  'Disponible
      .Text = Format(vDisponible, "Standard")
      
      .Col = 11  'Sobrante
      .Text = Format(vSobrante, "Standard")
       
     .MaxRows = .MaxRows + 1
     rs.MoveNext
    Loop
    rs.Close
    
End With

End Sub

Public Sub BoletaFOSOL(Expediente As Integer)
Dim strRuta As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strRuta = SIFGlobal.fxSIFPathReportes("FSL_BoletaExpdiente.rpt")

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Boleta Registro Expediente"
 .ReportFileName = strRuta
 
 .Connect = glogon.ConectRPT
 
 .SelectionFormula = "{FSL_EXPEDIENTES.COD_EXPEDIENTE}=" & Expediente
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

 .PrintReport
 
End With

Me.MousePointer = vbDefault

End Sub

Private Sub sbTraeResolucion(Expediente As Integer)
Dim i As Integer
Dim vKey As String, vLvw As MSComctlLib.ListView, vItem As MSComctlLib.ListItem
On Error GoTo vError

Me.lswComite.ColumnHeaders.Clear
Me.lswComite.ListItems.Clear
   
Set vLvw = Me.lswComite
    vLvw.ColumnHeaders.Add , , "Cédula", 1000
    vLvw.ColumnHeaders.Add , , "Nombre", 2500, 2

    
    strSQL = "select RESOLUCION_NOTAS,RESOLUCION_FECHA,FECHA_ESTABLECE_CAUSA" _
           & " from FSL_EXPEDIENTES " _
           & "where COD_EXPEDIENTE = '" & txtExpediente.Text & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    
    dtpFechaResolucion.Value = IIf(IsNull(rs!Resolucion_Fecha), vFecha, rs!Resolucion_Fecha)
    dtpFechaCausa.Value = IIf(IsNull(rs!Fecha_Establece_Causa), vFecha, rs!Fecha_Establece_Causa)
    txtResolucionNotas.Text = IIf(IsNull(rs!Resolucion_Notas), "", rs!Resolucion_Notas)
     
    rs.Close
    
    
    strSQL = "Select EC.COD_MIEMBRO,C.Cedula,C.Nombre " _
           & " from FSL_EXPEDIENTE_COMITE EC " _
           & " inner join FSL_COMITE_MIEMBROS C on EC.COD_MIEMBRO = C.COD_MIEMBRO" _
           & " Where EC.Cod_Expediente = '" & txtExpediente.Text & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
       
    Do While Not rs.EOF
      vKey = Trim(rs.Fields("COD_MIEMBRO")) & "(CA)"
      Set vItem = lswComite.ListItems.Add(, vKey, Trim(rs!Cedula))
                  vItem.SubItems(1) = rs!Nombre

      rs.MoveNext
    Loop
    
    rs.Close
    
Exit Sub
vError:
   MsgBox Err.Description, vbCritical
End Sub

Private Sub sbTraerHistoricoSeguimiento()
Dim i As Integer
On Error GoTo vError

With vGridHistoricoGestiones
 .MaxRows = 1
 .Row = .MaxRows
 For i = 1 To .MaxCols
   .Col = i
   .Text = ""
   .Tag = ""
 Next i

strSQL = "Select EG.CONSECUTIVO, EG.COD_EXPEDIENTE, EG.COD_GESTION,G.DESCRIPCION, EG.OBSERVACIONES " _
       & ", EG.USUARIO_REGISTRA , EG.FECHA_REGISTRO " _
       & " from FSL_EXPEDIENTE_GESTIONES EG " _
       & " inner join FSL_GESTIONES G on EG.COD_GESTION = G.COD_GESTION " _
       & " Where EG.COD_EXPEDIENTE = " & txtExpediente.Text & ""
rs.Open strSQL, glogon.Conection, adOpenForwardOnly


 Do While Not rs.EOF
  .Row = .MaxRows
  .Col = 1
  .Text = rs!FECHA_REGISTRO
      
  .Col = 2
  .Text = rs!Descripcion
  
  .Col = 3
  .AutoSize = True
  .Text = rs!observaciones
  
  .Col = 4
  .Text = rs!USUARIO_REGISTRA
      
  rs.MoveNext
  .MaxRows = .MaxRows + 1
 Loop
End With
rs.Close
Exit Sub
vError:
   MsgBox Err.Description, vbExclamation
End Sub


