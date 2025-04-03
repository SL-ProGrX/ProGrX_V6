VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmFSL_Expedientes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   13455
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
      TabIndex        =   36
      Top             =   1560
      Width           =   2775
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
      TabIndex        =   21
      Top             =   1560
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
      TabIndex        =   20
      Top             =   1920
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
      Left            =   4560
      MaxLength       =   15
      TabIndex        =   17
      ToolTipText     =   "Cédula Solicitante"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox cboParentesco 
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
      ItemData        =   "frmFSL_IngresoExpedientes.frx":0000
      Left            =   10320
      List            =   "frmFSL_IngresoExpedientes.frx":0019
      TabIndex        =   10
      Text            =   "cboParentesco"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contacto"
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
         TabIndex        =   11
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
      Width           =   1935
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
      TabPicture(0)   =   "frmFSL_IngresoExpedientes.frx":0062
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vgCreditos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle Causa"
      TabPicture(1)   =   "frmFSL_IngresoExpedientes.frx":007E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtObservaciones"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtEnfermedad"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Requisitos"
      TabPicture(2)   =   "frmFSL_IngresoExpedientes.frx":009A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lsw"
      Tab(2).Control(1)=   "lblEstadoRequisitos"
      Tab(2).Control(2)=   "Label14"
      Tab(2).ControlCount=   3
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
         TabIndex        =   27
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
         TabIndex        =   25
         Top             =   1020
         Width           =   10335
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   3615
         Left            =   -70680
         TabIndex        =   19
         Top             =   420
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   6376
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
      Begin FPSpreadADO.fpSpread vgCreditos 
         Height          =   4095
         Left            =   120
         TabIndex        =   24
         Top             =   420
         Width           =   12975
         _Version        =   524288
         _ExtentX        =   22886
         _ExtentY        =   7223
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   13
         SpreadDesigner  =   "frmFSL_IngresoExpedientes.frx":00B6
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label lblEstadoRequisitos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74400
         TabIndex        =   34
         Top             =   780
         Width           =   3375
      End
      Begin VB.Label Label14 
         Caption         =   "Requisitos"
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
         Left            =   -74400
         TabIndex        =   33
         Top             =   420
         Width           =   1095
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
         TabIndex        =   28
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
         TabIndex        =   26
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   116785155
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
            Picture         =   "frmFSL_IngresoExpedientes.frx":1A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":82A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":EB08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":1536A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":1BBCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":1BCE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":22548
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
            Picture         =   "frmFSL_IngresoExpedientes.frx":28DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":2F60C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":35E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":3C6D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":42F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":49794
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":5E906
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1296
      BandCount       =   4
      _CBWidth        =   13455
      _CBHeight       =   735
      _Version        =   "6.7.9816"
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
      Width3          =   1470
      NewRow3         =   0   'False
      Child4          =   "tlbMantenimiento"
      MinHeight4      =   330
      Width4          =   4095
      NewRow4         =   -1  'True
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
         TabIndex        =   16
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   30
         Width           =   2145
      End
      Begin MSComctlLib.Toolbar tlbMantenimiento 
         Height          =   330
         Left            =   165
         TabIndex        =   15
         Top             =   375
         Width           =   13200
         _ExtentX        =   23283
         _ExtentY        =   582
         ButtonWidth     =   1931
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
               Caption         =   "Borrar"
               Key             =   "Borrar"
               Object.ToolTipText     =   "Borrar Expediente"
               ImageIndex      =   5
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
         TabIndex        =   14
         Top             =   30
         Width           =   5265
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
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
         Left            =   8205
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   13
         ToolTipText     =   "Exposición a Riesgo de la persona"
         Top             =   30
         Width           =   5160
      End
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   7200
      Top             =   360
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
            Picture         =   "frmFSL_IngresoExpedientes.frx":752C8
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":8A43A
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":9F5AC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":B471E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":C9890
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":CA16A
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":DF2DC
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F444E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F4D28
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5602
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5720
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5846
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5970
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5B99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5DD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F5EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F600A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_IngresoExpedientes.frx":F6133
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFechaCausa 
      Height          =   315
      Left            =   10320
      TabIndex        =   29
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
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
      Format          =   116785155
      CurrentDate     =   40640
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2160
      TabIndex        =   31
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
      Format          =   116785155
      CurrentDate     =   40640
   End
   Begin VB.Label Label4 
      Caption         =   "Cédula Beneficiario"
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
      TabIndex        =   37
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image imgTags 
      Height          =   240
      Left            =   3840
      Picture         =   "frmFSL_IngresoExpedientes.frx":F651B
      ToolTipText     =   "Etiquetas"
      Top             =   840
      Width           =   240
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
      Height          =   255
      Left            =   8280
      TabIndex        =   35
      Top             =   840
      Width           =   5055
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
      TabIndex        =   32
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Left            =   8280
      TabIndex        =   30
      Top             =   1200
      Width           =   1935
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
      Left            =   8640
      TabIndex        =   23
      Top             =   1560
      Width           =   1575
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
      Left            =   9600
      TabIndex        =   22
      Top             =   1920
      Width           =   615
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
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image imgNo 
      Height          =   255
      Left            =   360
      Picture         =   "frmFSL_IngresoExpedientes.frx":F6625
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSi 
      Height          =   255
      Left            =   0
      Picture         =   "frmFSL_IngresoExpedientes.frx":F673F
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "Parentesco"
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
      Left            =   9360
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Nombre Beneficiario"
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
      Width           =   1095
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

Dim strSql As String
Dim rs As New ADODB.Recordset
Dim vFecha As String, vMembresiaMeses As Integer
Dim vRequisitosCompletos As Integer, vIdSolicitud As Long
Dim vTipoAplicacion As String, vEstadoRequisitos As String
Dim vPorcentajeSug As Double, vPorcentajeApl As Double
Dim vSaldo As Currency, vMORA_INT_COR As Currency, vMORA_INT_MOR As Currency, vMONTO_FOSOL As Currency
Dim vMonto As Currency

Private Sub cboTipo_Click()
  Call sbCargaCausas
End Sub

Private Sub Form_Activate()
 vModulo = 1
End Sub

Private Sub sbCargaRequisitos()
On Error GoTo vError

Dim strSql As String
Dim rs As New ADODB.Recordset
Dim vItem As MSComctlLib.ListItem
Dim vLvw As MSComctlLib.ListView
Dim vKey As String

Me.lsw.ColumnHeaders.Clear
Me.lsw.ListItems.Clear
  
If cboCausa.Text = Empty Then Exit Sub
  
Set vLvw = Me.lsw
    vLvw.ColumnHeaders.Add , , "Requisito", 2500
    vLvw.ColumnHeaders.Add , , "Opcional", 1000, 2
    vLvw.ColumnHeaders.Add , , "Asignado", 1000, 2

strSql = "Select distinct(RC.COD_REQUISITO), R.DESCRIPCION,RC.OPCIONAL,RC.ASIGNADO,ER.ESTADO as 'Presentado' " _
       & "from FSL_REQUISITOS_CAUSAS RC " _
       & " inner join FSL_REQUISITOS R on R.COD_REQUISITO = RC.COD_REQUISITO " _
       & " left join FSL_EXPEDIENTES_REQUISITOS ER on RC.COD_REQUISITO = ER.COD_REQUISITO " _
       & " Where RC.COD_CAUSA=" & SIFGlobal.fxSIFCodText(cboCausa) & " and RC.ASIGNADO <> 0 and " _
       & " ER.COD_EXPEDIENTE = " & txtExpediente.Text & ""

rs.Open strSql, glogon.Conection, adOpenStatic

If rs.EOF Then
  MsgBox "No se tiene los requisitos configurados de esta Causa de Liquidación"
  rs.Close
  Exit Sub
End If

Do While Not rs.EOF
  vKey = Trim(rs.Fields("COD_REQUISITO")) & "(CA)"
    
  Set vItem = lsw.ListItems.Add(, vKey, Trim(rs.Fields!Descripcion))
              vItem.SubItems(1) = IIf(rs!Opcional = 1, "Si", "No")
              vItem.SubItems(2) = IIf(rs!asignado = 1, "Si", "No")
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
  vModulo = 1
  vFecha = fxFechaServidor
  dtpFechaExpediente.Value = vFecha
  dtpFechaCausa.Value = vFecha
  
  Call sbCargaCombos
  Call sbCargaTipos
  Call sbCargaCausas
  SSTab.Tab = 0
  vgCreditos.MaxRows = 0
End Sub

Private Sub imgTags_Click()
If txtExpediente.Text <> Empty Then
   GLOBALES.gTag = txtExpediente.Text
   GLOBALES.gTag2 = txtCedula.Text
   Call sbSIFForms("frmFSL_SeguimientoEtiquetas", 1, , , False, Me)
Else
   MsgBox "Debe de registrar el Expediente", vbInformation
End If
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
   Select Case SSTab.Tab
      Case 2
        Call sbCargaRequisitos
   End Select
End Sub

Private Sub tlbMantenimiento_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Nuevo"
     Call sbLimpiar
     
  Case "Guardar"
     If txtCedula.Text = Empty Or txtNombre.Text = Empty Then
       MsgBox "Faltan los datos del Asociado"
       Exit Sub
     End If
     
     If fxValidaExpediente = False Then
       Call sbGuardaExpediente
       Call sbGuardaDetalleExpediente
       Call sbGuardaRequisitos(txtExpediente.Text)
     Else
       Call sbModificaExpediente
       Call sbGuardaDetalleExpediente
       Call sbGuardaRequisitos(txtExpediente.Text)
     End If
     
   Case "Imprimir"
     If txtExpediente.Text = Empty Then Exit Sub
     Call ReporteBoletaFOSOL(txtExpediente.Text)
     
   Case "Gestiones"
     If txtCedula.Text = Empty Or txtExpediente.Text = Empty Then Exit Sub
     GLOBALES.gTag = txtCedula.Text
     GLOBALES.gTag2 = txtNombre.Text
     GLOBALES.gTag3 = txtExpediente.Text
     Call sbSIFForms("frmFSL_Seguimiento")
   
   Case "Datos"
     Call sbSIFForms("frmAF_Principal")
     
       
End Select
  
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSql As String, rs As New ADODB.Recordset
Dim vCedTemp As String

On Error GoTo vError

'Busca primer en el Maestro de Socios, de lo contrario revisa si es una operacion
' y regresa la cedula de la operacion
vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 strSql = "select coalesce(count(*),0) as Existe from socios where cedula = '" & txtCedula & "'"
 rs.Open strSql, glogon.Conection, adOpenStatic
 If rs!Existe = 0 Then
   rs.Close
   strSql = "select cedula from reg_creditos where id_solicitud = " & txtCedula
   rs.Open strSql, glogon.Conection, adOpenStatic
   If Not rs.EOF And Not rs.BOF Then
      vCedTemp = Trim(rs!Cedula)
   End If
 End If
 rs.Close
 
    If vCedTemp = "" Then
     Call sbConsulta(txtCedula.Text)
    Else
     Call sbConsulta(vCedTemp)
    End If
Call sbTraeExpediente
'Trae los datos Aprobados por el cómite
Call sbTraeDatosAprobacion
End If

If KeyCode = vbKeyF4 Then Call sbBusqueda

Exit Sub
vError:
    MsgBox Err.Description

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
        Call sbConsulta(txtCedula)
    End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub txtExpediente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbTraeExpediente
  cboTipo.SetFocus
End If

If KeyCode = vbKeyF4 Then
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

strSql = "Select COD_EXPEDIENTE, COD_TIPO, COD_CAUSA, CEDULA, DOCUMENTO_REF,NUMERO_DOC_REF, PRESENTA_CEDULA" _
       & ", PRESENTA_NOMBRE, PRESENTA_CONTACTO,PRESENTA_PARENTESCO,MEMBRESIA_MESES, MEMBRESIA_PORCENTAJE, REQUISTOS_COMPLETOS" _
       & ", OBSERVACIONES, DETALLE_ENFERMEDAD, ENFERMEDAD_USUARIO,FECHA_ESTABLECE_CAUSA" _
       & ", REGISTRO_FECHA, REGISTRO_USUARIO, MODIFICA_USUARIO, MODIFICA_FECHA, ESTADO, RESOLUCION_ESTADO" _
       & ", RESOLUCION_NOTAS, RESOLUCION_FECHA, RESOLUCION_USUARIO, TOTAL_FOSOL, TOTAL_SOBRANTE" _
       & " from FSL_EXPEDIENTES "

If txtCedula.Text = Empty Then
  strSql = strSql & " where COD_EXPEDIENTE = " & txtExpediente.Text & " "
Else
  strSql = strSql & " where CEDULA = '" & txtCedula.Text & "' "
End If

rs.Open strSql, glogon.Conection, adOpenStatic

Do While Not rs.EOF
  txtCedula.Text = rs!Cedula
  txtExpediente.Text = rs!Cod_Expediente
  cboDocumento.Text = rs!DOCUMENTO_REF
  txtNumDocumento.Text = rs!NUMERO_DOC_REF
  dtpFechaExpediente.Value = rs!Registro_Fecha
  txtPresentaCedula.Text = rs!PRESENTA_CEDULA
  txtPresentanombre.Text = rs!PRESENTA_NOMBRE
  txtPresentaContacto.Text = rs!PRESENTA_CONTACTO
  cboTipo.Text = fxTipo(rs!COD_TIPO)
  cboCausa.Text = fxCausa(rs!COD_CAUSA)
  cboParentesco.Text = rs!PRESENTA_PARENTESCO
  txtObservaciones.Text = rs!observaciones
  txtEnfermedad.Text = IIf(IsNull(rs!DETALLE_ENFERMEDAD), "No se especifica", rs!DETALLE_ENFERMEDAD)
  dtpFechaCausa.Value = rs!FECHA_ESTABLECE_CAUSA
  Select Case rs!Estado
    Case "R"
       txtEstado.Text = "Recibida"
    Case "A"
       txtEstado.Text = "Aprobada"
    Case "P"
       txtEstado.Text = "Pendiente"
  End Select
  rs.MoveNext
Loop
rs.Close

Call sbConsulta(txtCedula.Text)

Exit Sub
vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Function fxTipo(vCodigo As String) As String
Dim Record As New ADODB.Recordset
Dim strSql As String

strSql = "Select COD_TIPO, DESCRIPCION from FSL_TIPOS where COD_TIPO='" & vCodigo & "'"
Record.Open strSql, glogon.Conection, adOpenStatic

fxTipo = CStr(Record!COD_TIPO & "-" & Record!Descripcion)

Record.Close

End Function

Private Function fxCausa(vCodigo As String) As String
Dim Record As New ADODB.Recordset
Dim strSql As String

strSql = "Select COD_CAUSA, DESCRIPCION from FSL_CAUSAS where COD_CAUSA='" & vCodigo & "'" _
       & " and COD_TIPO= " & SIFGlobal.fxSIFCodText(cboTipo) & ""
Record.Open strSql, glogon.Conection, adOpenStatic

fxCausa = CStr(Record!COD_CAUSA & "-" & Record!Descripcion)

Record.Close

End Function

Private Sub sbCargaCombos()
    cboParentesco.Clear
    cboParentesco.AddItem "Padre"
    cboParentesco.AddItem "Madre"
    cboParentesco.AddItem "Hijos(a)"
    cboParentesco.AddItem "Conyugue"
    cboParentesco.AddItem "Otro"
    cboParentesco.Text = "Otro"
    
    cboDocumento.Clear
    cboDocumento.AddItem "Accion de personal"
    cboDocumento.AddItem "Acta defunsion"
    cboDocumento.Text = "Accion de personal"
End Sub

Private Sub sbCargaTipos()
On Error GoTo vError
  
strSql = "select COD_TIPO,DESCRIPCION From FSL_TIPOS Where ACTIVO = 1"
rs.Open strSql, glogon.Conection, adOpenStatic

Do While Not rs.EOF
   cboTipo.AddItem (rs!COD_TIPO & "-" & Trim(rs!Descripcion))
   rs.MoveNext
Loop
  
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboTipo.Text = rs!COD_TIPO & " - " & rs!Descripcion
Else
  MsgBox "Falta registro de Tipos de Liquidación y Causas!!!"
End If

rs.Close

Exit Sub
vError:
  MsgBox Err.Description
End Sub

Private Sub sbCargaCausas()
On Error GoTo vError
   
If cboTipo.Text = Empty Then Exit Sub
   
cboCausa.Clear
strSql = "Select COD_CAUSA,DESCRIPCION from FSL_CAUSAS " _
       & " where COD_TIPO = " & SIFGlobal.fxSIFCodText(cboTipo) & ""
rs.Open strSql, glogon.Conection, adOpenStatic

Do While Not rs.EOF
   cboCausa.AddItem (rs!COD_CAUSA & "-" & Trim(rs!Descripcion))
   rs.MoveNext
Loop

If rs.RecordCount > 0 Then
  rs.MoveFirst
  cboCausa.Text = rs!COD_CAUSA & " - " & rs!Descripcion
End If

rs.Close
      
Exit Sub
vError:
   MsgBox Err.Description
End Sub

Private Sub sbGuardaExpediente()

strSql = "INSERT FSL_EXPEDIENTES (COD_EXPEDIENTE, COD_TIPO, COD_CAUSA, CEDULA, DOCUMENTO_REF,NUMERO_DOC_REF, PRESENTA_CEDULA, PRESENTA_NOMBRE" _
       & ", PRESENTA_CONTACTO,PRESENTA_PARENTESCO,FECHA_ESTABLECE_CAUSA, MEMBRESIA_MESES, MEMBRESIA_PORCENTAJE, REQUISTOS_COMPLETOS, OBSERVACIONES, LIQUIDACION_ESTADO, LIQUIDACION_NUMERO" _
       & ", REGISTRO_FECHA, REGISTRO_USUARIO, ESTADO,DETALLE_ENFERMEDAD)" _
       & "  Values (" & txtExpediente & ",'" & SIFGlobal.fxSIFCodText(cboTipo) & "','" & SIFGlobal.fxSIFCodText(cboCausa) & "','" & txtCedula & "'" _
       & ", '" & cboDocumento.Text & "','" & txtNumDocumento.Text & "','" & txtPresentaCedula.Text & "','" & txtPresentanombre.Text & "'" _
       & ", '" & txtPresentaContacto & "','" & cboParentesco.Text & "','" & Format(dtpFechaCausa.Value, "yyyymmdd") & "'," & vMembresiaMeses & ",0,'" & vRequisitosCompletos & "','" & txtObservaciones.Text & "','',''" _
       & ", '" & Format(vFecha, "yyyymmdd") & "','" & glogon.Usuario & "','R','" & txtEnfermedad.Text & "')"
       
glogon.Conection.Execute strSql

End Sub

Private Sub sbModificaExpediente()
strSql = "UPDATE FSL_EXPEDIENTES SET DOCUMENTO_REF ='" & cboDocumento.Text & "' ,NUMERO_DOC_REF ='" & txtNumDocumento.Text & "',REGISTRO_FECHA = '" & Format(dtpFechaExpediente, "yyyymmdd") & "'" _
       & ", COD_TIPO ='" & SIFGlobal.fxSIFCodText(cboTipo) & "' ,COD_CAUSA ='" & SIFGlobal.fxSIFCodText(cboCausa) & "' ,FECHA_ESTABLECE_CAUSA = '" & Format(dtpFechaCausa, "yyyymmdd") & "'  " _
       & ", ESTADO = 'R',REQUISTOS_COMPLETOS ='" & vRequisitosCompletos & "',PRESENTA_CEDULA = '" & txtPresentaCedula.Text & "',PRESENTA_NOMBRE ='" & txtPresentanombre.Text & "'" _
       & ", PRESENTA_CONTACTO = '" & txtPresentaContacto & "',PRESENTA_PARENTESCO = '" & cboParentesco.Text & "'" _
       & ", OBSERVACIONES ='" & txtObservaciones.Text & "' ,MODIFICA_USUARIO ='" & glogon.Usuario & "' " _
       & ", MODIFICA_FECHA = '" & Format(vFecha, "yyyymmdd") & "',DETALLE_ENFERMEDAD='" & UCase(txtEnfermedad.Text) & "'" _
       & " WHERE  CEDULA = '" & Trim(txtCedula.Text) & "'"
       
glogon.Conection.Execute strSql
End Sub

'Guarda el detalle de la aplicacion
Private Sub sbGuardaDetalleExpediente()
Dim i As Integer, vExiste As Boolean

vPorcentajeApl = 0
vMonto = 0
vSaldo = 0
vPorcentajeSug = 0
vMONTO_FOSOL = 0
vExiste = True

With vgCreditos

For i = 1 To .MaxRows
  .Row = i
    
  .Col = 2
  If .Value = 1 Then
    .Col = 1
    vPorcentajeApl = IIf(Format(.Text, "standard") = Empty, 0, Format(.Text, "standard"))
    .Col = 2
    vIdSolicitud = .CellTag
    .Col = 4
    vMonto = Format(.Text, "standard")
    .Col = 5
    vSaldo = Format(.Text, "standard")
    .Col = 9
    vPorcentajeSug = Format(.Text, "standard")
    .Col = 13
    .Text = CCur(Format(vMONTO_FOSOL, "standard"))
    
    strSql = "Select COD_EXPEDIENTE,ID_SOLICITUD from FSL_EXPEDIENTES_DETALLE" _
           & " where COD_EXPEDIENTE = " & txtExpediente.Text & " and ID_SOLICITUD = " & vIdSolicitud & ""
    rs.Open strSql, glogon.Conection, adOpenStatic
         
    If rs.EOF Then
      vExiste = False
    End If
    
    rs.Close
    
    If vExiste Then
       strSql = "Update FSL_EXPEDIENTES_DETALLE set PORCENTAJE_SUGERIDO = " & vPorcentajeSug & ",PORCENTAJE_APLICADO = " & vPorcentajeApl & "" _
              & ", ACTUALIZACION_FECHA = '" & Format(vFecha, "yyyymmdd") & "',MODIFICA_USUARIO = '" & glogon.Usuario & "'" _
              & ", MODIFICA_FECHA= '" & Format(vFecha, "yyyymmdd") & "', TIPO_APLICACION_FOSOL = '" & vTipoAplicacion & "',MONTO_FOSOL = " & vMONTO_FOSOL & "" _
              & " where COD_EXPEDIENTE = " & txtExpediente.Text & " and ID_SOLICITUD = " & vIdSolicitud & ""
      glogon.Conection.Execute strSql
    Else
      strSql = "INSERT FSL_EXPEDIENTES_DETALLE (COD_EXPEDIENTE,ID_SOLICITUD,SALDO,MORA_INT_COR" _
             & ", MORA_INT_MOR,MORA_CARGOS,MONTO_FORMALIZADO,TOTAL_DEUDA,PORCENTAJE_SUGERIDO,PORCENTAJE_APLICADO" _
             & ", MONTO_FOSOL,ACTUALIZACION_FECHA,MODIFICA_USUARIO,MODIFICA_FECHA,TIPO_APLICACION_FOSOL)" _
             & " Values (" & txtExpediente.Text & "," & vIdSolicitud & "," & vSaldo & ",'0','0','0'," & vMonto & "," & vSaldo & "" _
             & ", " & vPorcentajeSug & "," & vPorcentajeApl & "," & vMONTO_FOSOL & ",'" & Format(vFecha, "yyyymmdd") & "'" _
             & ", '" & glogon.Usuario & "','" & Format(vFecha, "yyyymmdd") & "','" & vTipoAplicacion & "')"
      glogon.Conection.Execute strSql
    End If 'Existe= true
 End If 'value=0
Next i

End With
End Sub

Private Sub sbGuardaRequisitos(vExpediente As Integer)
Dim vExiste As Boolean, vCodRequisito As String
Dim i As Integer, vOpcional As Integer

With lsw
For i = 1 To .ListItems.Count
  vCodRequisito = DeCodificaPrimaryKey(.ListItems.Item(i).Key, 1, "(CA)")
  strSql = "select count(COD_REQUISITO) as 'Requisito' from FSL_EXPEDIENTES_REQUISITOS" _
         & " where COD_EXPEDIENTE = " & vExpediente & " and COD_REQUISITO = " & vCodRequisito & "  "
  rs.Open strSql, glogon.Conection, adOpenStatic
      
  If rs!Requisito >= 0 Then
     vExiste = True
  End If
  rs.Close
  
  If .ListItems.Item(i).Checked Then
    If .ListItems(i).SubItems(1) = "Si" Then
        vOpcional = 1
    Else
        vOpcional = 0
    End If
      
    If vExiste = False Then
      strSql = "INSERT FSL_EXPEDIENTES_REQUISITOS (COD_EXPEDIENTE,COD_REQUISITO,OPCIONAL,ESTADO,REGISTRO_FECHA,REGISTRO_USUARIO)" _
             & "Values(" & vExpediente & ",'" & vCodRequisito & "'," & vOpcional & ",1,'" & Format(vFecha, "yyyymmdd") & "','" & glogon.Usuario & "')"
      glogon.Conection.Execute strSql
    Else
      strSql = "UPDATE FSL_EXPEDIENTES_REQUISITOS SET OPCIONAL = " & vOpcional & ",ESTADO = 1,REGISTRO_FECHA = '" & Format(vFecha, "yyyymmdd") & "',REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
             & "WHERE COD_EXPEDIENTE=" & vExpediente & " and COD_REQUISITO='" & vCodRequisito & "'"
      glogon.Conection.Execute strSql
    End If 'vExiste = False
  Else '.ListItems.Item(i).Checked
    If vExiste = True Then
      strSql = "UPDATE FSL_EXPEDIENTES_REQUISITOS SET OPCIONAL = " & vOpcional & ",ESTADO = 0,REGISTRO_FECHA = '" & Format(vFecha, "yyyymmdd") & "',REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
             & "WHERE COD_EXPEDIENTE=" & vExpediente & " and COD_REQUISITO='" & vCodRequisito & "'"
      glogon.Conection.Execute strSql
    End If 'vExiste = False

  End If ' no(.ListItems.Item(1).Checked)
Next i
End With
End Sub

'Verdadero si Existe el expediente
Private Function fxValidaExpediente() As Boolean
On Error GoTo vError
    
If txtExpediente.Text = Empty Then
  strSql = "Select isnull(max(COD_EXPEDIENTE) + 1, 1) as 'Expediente' from FSL_EXPEDIENTES"
  rs.Open strSql, glogon.Conection, adOpenStatic
   txtExpediente.Text = rs!Expediente
  rs.Close
  fxValidaExpediente = False
Else
  fxValidaExpediente = True
End If

If fxRequisitosCompletos = True Then
   vRequisitosCompletos = 1
   lblEstadoRequisitos.Caption = "Requisitos Básicos Completos"
   lblEstadoRequisitos.ForeColor = vbBlue
Else
   vRequisitosCompletos = 0
   lblEstadoRequisitos.Caption = "Faltan requisitos Básicos por presentar"
   lblEstadoRequisitos.ForeColor = vbBlue
   MsgBox "Faltan requisitos para la aprobación", vbExclamation
End If

txtPresentaCedula.Text = IIf(txtPresentaCedula.Text = Empty, txtCedula.Text, txtPresentaCedula.Text)
txtPresentanombre.Text = IIf(txtPresentanombre.Text = Empty, txtNombre.Text, txtPresentanombre.Text)
txtPresentaContacto.Text = IIf(txtPresentaContacto.Text = Empty, "Sin datos del Beneficiario", txtPresentaContacto.Text)
txtEnfermedad.Text = IIf(txtEnfermedad.Text = Empty, "Sin definicion", txtEnfermedad.Text)
txtObservaciones.Text = IIf(txtObservaciones.Text = Empty, "Sin Observaciones registradas", txtObservaciones.Text)

Exit Function
vError:
   MsgBox Err.Description
End Function

Private Sub sbLimpiar()
Dim Control As Control

For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
      Control = Empty
    ElseIf TypeOf Control Is ComboBox Then
      Control.Clear
    ElseIf TypeOf Control Is fpSpread Then
      Control.MaxRows = 0
    End If
Next

End Sub

Private Sub sbConsulta(pCedula As String)
Dim strSql As String, rs As New ADODB.Recordset
Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset
Dim vEstadoSocio As String
     
vFianzas = False
    
strSql = "select S.cedula as CedulaX,S.nombre,S.fechaingreso,S.estadoactual,S.notas,S.bloqueo,S.nota_User,S.nota_Fecha,A.*" _
       & ",dbo.fxCRDClasificacion(S.cedula,getdate()) as Clasificacion,dbo.fxSIFRatePersona(S.cedula) as Rating" _
       & ",dbo.fxSIFMensajesNumero(S.cedula) as IndMensajes, dbo.fxCBRHistorialNumero(S.cedula) as IndCobro" _
       & ",dbo.fxCBRFianzasEnMora(S.cedula) as IndFianzas" _
       & ",I.descripcion as InstitucionX, E.descripcion as EstadoX" _
       & " from socios S left join Ahorro_consolidado A on S.cedula = A.cedula" _
       & " inner join afi_estados_persona E on S.estadoActual = E.cod_estado" _
       & " inner join instituciones I on S.cod_institucion = I.cod_Institucion" _
       & " Where S.cedula = '" & Trim(pCedula) & "'"

rs.CursorLocation = adUseServer
rs.Open strSql, glogon.Conection, adOpenStatic
 
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
      strSql = "select count(*) as Existe from afi_cr_renuncias where cedula = '" & pCedula & "' and estado = 'T'"
      rsTmp.Open strSql, glogon.Conection, adOpenStatic
      If rsTmp!Existe > 0 Then
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
       
   SSTab.Tab = 0
   
   'Actualiza el Detalle de Creditos
   Call sbCreditos
 
 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
   
   
 
End Sub

Private Sub sbCreditos(Optional pSheet As Integer = 1)
Dim rs As New ADODB.Recordset, strSql As String
Dim curCuota As Currency, curMonto As Currency
Dim curSaldo As Currency, vMora As Boolean
Dim i As Integer

'On Error Resume Next

curCuota = 0
curMonto = 0
curSaldo = 0
vMora = False

Me.MousePointer = vbHourglass
vMora = False

With vgCreditos
 .Sheet = pSheet
 .ActiveSheet = pSheet
 
 .MaxRows = 0
 strSql = "exec spFSLCreditos '" & txtCedula.Text & "','" & Mid(.SheetName, 1, 1) & "'"
 
 rs.CursorLocation = adUseServer
 rs.Open strSql, glogon.Conection, adOpenStatic

Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows

vTipoAplicacion = rs!TipoAplicacion

For i = 1 To .MaxCols
  .Col = i
  Select Case i
    Case 1 'Status
          .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
    
         Select Case rs!ProcesoCod
          Case "N"
   
            If Not IsNull(rs!referencia) Then
                If rs!MoraCuota = 0 Then
                  .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
                  .TextTip = TextTipFixed
                  .TextTipDelay = 1000
                  .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
                  .CellNoteIndicatorColor = vbRed
                  .CellNote = "Referencia: " & rs!referencia
                End If
                .FontBold = True
            End If
    
            If rs!IndicadorCbr > 0 Then
              .TypePictPicture = imgSemaforos.ListImages.Item(9).Picture
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
            
              .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
              .CellNoteIndicatorColor = vbRed
              
              .CellNote = "!!! Esta Operación fue Reversada de Cobro Judicial, Revise el Tab de Cobros para mayor información..!!!"
                        
            End If
          
          Case "J"
              .TypePictPicture = imgSemaforos.ListImages.Item(7).Picture
               vMora = True
                   
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
            
              .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
              .CellNoteIndicatorColor = vbRed
              
              .CellNote = ">> Cobro Judicial <<" & vbCrLf _
                        & "Fecha : " & Format(rs!fecha_enviaproceso, "dd/mm/yyyy") & vbCrLf _
                        & "Nota  : " & rs!observacion_proceso & ""
          
          Case "T"
                If rs!MoraCuota = 0 Then .TypePictPicture = imgSemaforos.ListImages.Item(10).Picture
                
                If rs!IndicadorCbr > 0 Then
                   .TypePictPicture = imgSemaforos.ListImages.Item(9).Picture
                End If
    
         End Select
         
         If Mid(rs!Estado, 1, 1) = "C" Then
            .TypePictPicture = imgSemaforos.ListImages.Item(6).Picture
         End If

        ' Si esta moroso indicar Mora siempre y cuando no este en cobro Judicial
        If rs!MoraCuota > 0 And rs!ProcesoCod <> "J" Then
          
          .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
          vMora = True
        
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
        
          .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
          .CellNoteIndicatorColor = vbBlue
          
          .CellNote = "Referencia..:" & rs!referencia & vbCrLf & "Morosidad:  Cuotas: " & rs!MoraCuota & vbCrLf _
                    & "   Intereses : " & Format(rs!MoraInt, "Standard") & vbCrLf _
                    & "   Cargos    : " & Format(rs!MoraCargos, "Standard") & vbCrLf _
                    & "   Póliza    : " & Format(rs!MoraPoliza, "Standard") & vbCrLf _
                    & "   Principal : " & Format(rs!MoraPrincipal, "Standard") & vbCrLf _
                    & "   Cta.+ Vieja : " & Format(rs!MoraAntigua, "####-##") & vbCrLf _
                    & "   Cta. Ultima : " & Format(rs!MoraUltima, "####-##") & vbCrLf & vbCrLf _
                    & "   Total Mora  : " & Format(rs!MoraInt + rs!MoraCargos + rs!MoraPrincipal + rs!MoraPoliza, "Standard") & vbCrLf
        
        End If
    
    Case 2 'Operacion
       .CellTag = CStr(rs!id_solicitud)
       .TypeCheckText = CStr(rs!id_solicitud)
            
    Case 3 'Primer Deduccion
        .Text = Format(rs!Prideduc, "####-##")
    Case 4 'Monto
        .Text = Format(rs!montoapr, "Standard")
    Case 5 'Saldo
        .Text = Format(rs!Saldo, "Standard")
    Case 6 'Cuota
        .Text = Format(rs!Cuota, "Standard")
    Case 7 'Ultimo Movimiento
        .Text = Format(rs!FecUlt, "####-##")
    Case 8 'Garantia
        .Text = rs!Garantia
        .TextTip = TextTipFixed
        .TextTipDelay = 1000
        .CellNote = rs!GarantiaDetalle
    Case 12 'Estado
        .Text = rs!Estado
    Case 13 'Proceso
        .Text = rs!Proceso
    
    .Col = 8
    If (.Text = "Hipotecaria") Or (.Text = "Polizas") Or (rs!Plazo >= 900) Or _
       (.Text = "Cuota Mantenimiento") Or (.Text = "Fondo Solidario") Or _
       (.Text = "Retencion") Then
      .Col = 2
      .Lock = True
    Else
      If vTipoAplicacion = "M" Then
         .Col = 10
         .Value = 1
      Else
         .Col = 11
         .Value = 1
      End If
      .Col = 9
      .Text = rs!PorcentajeAplicacion
    End If
    
    Case 15 'Liquidacion
        .Text = CStr(rs!LIQUIDACION)
    Case 16 'Fecha Liquidación
        .Text = rs!FECHALIQUIDACION
    
    
  End Select
Next i
rs.MoveNext
Loop
rs.Close
End With

Me.MousePointer = vbDefault

End Sub

Private Sub vgCreditos_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim i As Integer

Select Case NewSheet
    Case 1 'Activos
       Call sbCreditos(NewSheet)
       Call sbTraeDatosAprobacion

     Case 2 'Cancelados
       Call sbCreditos(NewSheet)
       Call sbTraeDatosAprobacion
End Select
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
  If lsw.ListItems(i).SubItems(2) = "Si" Then
    If Not (lsw.ListItems(i).Checked) Then
        fxRequisitosCompletos = False
    End If
  End If
Next i
End With

End Function

Private Sub sbTraeDatosAprobacion()
Dim i As Integer
With vgCreditos
   
If txtExpediente.Text = Empty Then Exit Sub
   
strSql = "Select COD_EXPEDIENTE, ID_SOLICITUD, SALDO, MORA_INT_COR, MORA_INT_MOR, MORA_CARGOS " _
       & ", TOTAL_DEUDA, MONTO_LIQUIDACION, PORCENTAJE_SUGERIDO, PORCENTAJE_APLICADO, MONTO_FOSOL" _
       & ", TIPO_APLICACION_FOSOL, MONTO_FORMALIZADO" _
       & " from FSL_EXPEDIENTES_DETALLE where COD_EXPEDIENTE = '" & txtExpediente.Text & "' "
rs.Open strSql, glogon.Conection, adOpenStatic

Do While Not rs.EOF
    For i = 1 To .MaxRows
      .Row = i
      .Col = 3
      If rs!id_solicitud = .CellTag Then
         .Col = 1
         .Text = CCur(rs!PORCENTAJE_APLICADO)
         .Col = 3
         .Value = 1
         If rs!TIPO_APLICACION_FOSOL = "M" Then
          .Col = 11
          .Value = 1
          .Col = 12
          .Value = 0
         Else
          .Col = 12
          .Value = 1
          .Col = 11
          .Value = 0
         End If
         .Col = 13
         .Text = Format(rs!MONTO_FOSOL, "standard")
      End If
    Next i
rs.MoveNext
Loop
rs.Close
End With
End Sub

Public Sub ReporteBoletaFOSOL(Expediente As Integer)
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

' .StoredProcParam(0) = "FRM"
' .StoredProcParam(1) = Expediente
' .StoredProcParam(2) = 0

 .Action = 1
 
End With

Me.MousePointer = vbDefault

End Sub


