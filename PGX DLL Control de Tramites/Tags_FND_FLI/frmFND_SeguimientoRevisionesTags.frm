VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmFNDSeguimientoRevisionesTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisión de Fondos (Tags)"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFND_SeguimientoRevisionesTags.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContrato 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtPlan 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   720
      Width           =   1455
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10560
      Top             =   0
   End
   Begin VB.TextBox txtCedula 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":6852
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":D0B4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":13916
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":1A178
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":209DA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2723C
            Key             =   "IMG6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   8520
      Top             =   0
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
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2DA9E
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2DBBC
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2DCE2
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2DE0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2DF1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2E035
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2E136
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2E26D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2E382
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   8760
      Top             =   0
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
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":2E4A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":34D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":3B56A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_SeguimientoRevisionesTags.frx":3B684
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Fondos"
      TabPicture(0)   =   "frmFND_SeguimientoRevisionesTags.frx":3B7A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgRefresh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tlbRefresh"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vGrid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFND_SeguimientoRevisionesTags.frx":3B7BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDescripcion"
      Tab(1).Control(1)=   "txtOperadora"
      Tab(1).Control(2)=   "lblFecha"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "lblEstado"
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(6)=   "lblContrato"
      Tab(1).Control(7)=   "Label1"
      Tab(1).Control(8)=   "lblOperacionASE"
      Tab(1).Control(9)=   "lblIncTipo"
      Tab(1).Control(10)=   "lblIncAnual"
      Tab(1).Control(11)=   "Label7(2)"
      Tab(1).Control(12)=   "lblRendimiento"
      Tab(1).Control(13)=   "lblAportes"
      Tab(1).Control(14)=   "lblRenueva"
      Tab(1).Control(15)=   "lblPlazo"
      Tab(1).Control(16)=   "Label7(6)"
      Tab(1).Control(17)=   "Label7(5)"
      Tab(1).Control(18)=   "Label7(4)"
      Tab(1).Control(19)=   "Label7(3)"
      Tab(1).Control(20)=   "Label7(1)"
      Tab(1).Control(21)=   "Label7(7)"
      Tab(1).Control(22)=   "lblTotal"
      Tab(1).Control(23)=   "lblMonto"
      Tab(1).Control(24)=   "Label7(0)"
      Tab(1).Control(25)=   "lblX"
      Tab(1).Control(26)=   "Label2(1)"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "Seguimiento"
      TabPicture(2)   =   "frmFND_SeguimientoRevisionesTags.frx":3B7DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vGridSeguimiento"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Revisión"
      TabPicture(3)   =   "frmFND_SeguimientoRevisionesTags.frx":3B7F6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboEtiquetas"
      Tab(3).Control(1)=   "txtObservacion"
      Tab(3).Control(2)=   "tlbAplicar"
      Tab(3).Control(3)=   "lswErrores"
      Tab(3).Control(4)=   "Label27"
      Tab(3).Control(5)=   "Label2(0)"
      Tab(3).Control(6)=   "Label8(1)"
      Tab(3).ControlCount=   7
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox txtOperadora 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   5895
      End
      Begin VB.ComboBox cboEtiquetas 
         Height          =   330
         ItemData        =   "frmFND_SeguimientoRevisionesTags.frx":3B812
         Left            =   -73320
         List            =   "frmFND_SeguimientoRevisionesTags.frx":3B814
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtObservacion 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   -73320
         MaxLength       =   995
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   4080
         Width           =   9135
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5415
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   10575
         _Version        =   524288
         _ExtentX        =   18653
         _ExtentY        =   9551
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
         MaxCols         =   7
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmFND_SeguimientoRevisionesTags.frx":3B816
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridSeguimiento 
         Height          =   5775
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   10575
         _Version        =   524288
         _ExtentX        =   18653
         _ExtentY        =   10186
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
         MaxCols         =   487
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmFND_SeguimientoRevisionesTags.frx":3C2B4
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   570
         Left            =   -65760
         TabIndex        =   9
         Top             =   5880
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1005
         ButtonWidth     =   2117
         ButtonHeight    =   1005
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicar Etiqueta"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswErrores 
         Height          =   2655
         Left            =   -73320
         TabIndex        =   10
         Top             =   1200
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Aplicado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Mensaje"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbRefresh 
         Height          =   330
         Left            =   9480
         TabIndex        =   43
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonWidth     =   1984
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgRefresh"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refrescar"
               Key             =   "Refrescar"
               Object.ToolTipText     =   "Volver a cargar la información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         MousePointer    =   1
      End
      Begin MSComctlLib.ImageList imgRefresh 
         Left            =   8760
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFND_SeguimientoRevisionesTags.frx":3C8AC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -69120
         TabIndex        =   40
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -69840
         TabIndex        =   39
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -71280
         TabIndex        =   38
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -72000
         TabIndex        =   37
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblContrato 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -73560
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contrato"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74760
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblOperacionASE 
         Caption         =   "Operación Asociada >>"
         Height          =   255
         Left            =   -74640
         TabIndex        =   34
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label lblIncTipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -71280
         TabIndex        =   33
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblIncAnual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -73560
         TabIndex        =   32
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inc. Anual"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   -74760
         TabIndex        =   31
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblRendimiento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -73560
         TabIndex        =   30
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblAportes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -69120
         TabIndex        =   29
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblRenueva 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -69120
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblPlazo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -71280
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Renueva"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   -69840
         TabIndex        =   26
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rendimiento"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   -74760
         TabIndex        =   25
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aportes"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   -69840
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inc.Tipo"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   -72000
         TabIndex        =   23
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plazo"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   -72000
         TabIndex        =   22
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   -72000
         TabIndex        =   21
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -71280
         TabIndex        =   20
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -73560
         TabIndex        =   19
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   -74760
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operadora"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74760
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plan"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   -74760
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Omisiones"
         Height          =   375
         Left            =   -74760
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Etiqueta"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Observación"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   11
         Top             =   4080
         Width           =   1455
      End
   End
   Begin VB.Label LblCedula 
      Caption         =   "Cedula"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "Nombre"
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblNombreUsuario 
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1440
      Picture         =   "frmFND_SeguimientoRevisionesTags.frx":3C9D1
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmFNDSeguimientoRevisionesTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperadora As Integer, vPlan As String, vContrato As String
Dim vPaso As Boolean



Private Sub sbCargarObservacion()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError
    
    strSQL = "select ISNULL(MENSAJE,'') from SIF_TAGS_AVISOS where TAG_CODIGO = '" & SIFGlobal.fxCodText(cboEtiquetas.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        txtObservacion = rs.Fields(0) & vbNewLine
    Else
        txtObservacion = Empty
    End If
    
    For i = 1 To lswErrores.ListItems.Count
        If lswErrores.ListItems(i).Checked = True Then
            If lswErrores.ListItems(i).SubItems(2) = "N" Then
                If txtObservacion = Empty Then
                    txtObservacion.Text = "-" & lswErrores.ListItems(i).SubItems(3)
                Else
                    txtObservacion.Text = txtObservacion.Text & vbNewLine & "-" & lswErrores.ListItems(i).SubItems(3)
                End If
            End If
        End If
    Next
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub cboEtiquetas_Click()
If vPaso Then Exit Sub
Call sbCargarObservacion
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
vModulo = 8

lblNombreUsuario.Caption = glogon.Usuario

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub lswErrores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    If Item.SubItems(2) = "S" Then
        Item.Checked = True
        If MsgBox("El error ya fué aplicado desea agregar únicamente la nota", vbOKCancel) = vbOK Then
              If txtObservacion = Empty Then
                txtObservacion.Text = " - " & Item.SubItems(1)
              Else
                txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
              End If
        End If
        Exit Sub
    End If
    
    If Item.Checked Then
    
      strSQL = "insert SIF_OMISIONESG (cedula,ID_ERROR,MODULO,CODIGO,DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO) values('" & txtCedula.Text _
             & "'," & Item.Text & ",'FND','" & txtPlan.Text & "','" & txtContrato.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
      Call ConectionExecute(strSQL)
             
      strSQL = "select max(LINEA_ERR) as 'Linea' from SIF_OMISIONESG where codigo = '" & txtPlan.Text & "' and Documento = '" & txtContrato.Text & "' and ID_ERROR = " & Item.Text
      Call OpenRecordSet(rs, strSQL)
          Item.Tag = rs!Linea
      rs.Close
      
      If txtObservacion = Empty Then
        txtObservacion.Text = " - " & Item.SubItems(1)
      Else
        txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
      End If
      
    Else
      strSQL = "delete SIF_OMISIONESG where LINEA_ERR = " & Item.Tag
      Call ConectionExecute(strSQL)
      Item.Tag = ""

      Call sbCargarObservacion
    End If
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
   Case 1
     
     Call sbCargaDetalle
   Case 2
     Call sbCargarGridSeguimiento
   Case 3
     Call sbCargarListaErrores
     Call sbCargarCombosEtiquetas
End Select
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbCargarListaFondos

End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

Me.MousePointer = vbHourglass

If Trim(cboEtiquetas.Text) = Empty Then
    MsgBox "Debe seleccionar la etiqueta que desea plicar"
    Me.MousePointer = vbDefault
    Exit Sub
End If

If MsgBox("Está seguro que sea aplicar la etiqueta en las afiliaciones seleccionadas", vbExclamation + vbYesNo) = vbNo Then
    Me.MousePointer = vbDefault
    Exit Sub
End If

Call sbSIFRegistraTags(vPlan, SIFGlobal.fxCodText(cboEtiquetas.Text), txtObservacion.Text, vContrato, "FND", vPlan, vContrato, txtCedula.Text)

Call sbAplicarErrores
Call sbCargarListaFondos

txtCedula.SetFocus
SSTab1.Tab = 0

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbAplicarErrores()
'' Procedimiento para colocar los errores ingresados en aplicados
Dim Linea As String, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    strSQL = "update SIF_OMISIONESG SET APLICADO = 'S' WHERE cedula = '" & txtCedula.Text _
           & "' AND MODULO = 'FND' AND CODIGO = '" & txtPlan.Text & "' AND DOCUMENTO = '" & txtContrato.Text & "'"
    Call ConectionExecute(strSQL)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargarListaFondos(Optional ByVal strCedula As String = Empty)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "SELECT 'B', F.CEDULA,S.nombre,F.usuario,F.COD_CONTRATO,F.COD_PLAN,F.COD_OPERADORA" _
       & " FROM FND_CONTRATOS F   inner join SOCIOS S on F.CEDULA = S.CEDULA" _
       & " LEFT JOIN SIF_OFICINAS O ON F.COD_OFICINA = O.COD_OFICINA" _
       & " WHERE isnull(F.ANALISTA_REVISION,'N') = 'N' and F.estado = 'A'"


If Trim(txtCedula.Text) <> "" Then
    strSQL = strSQL & " and F.Cedula = '" & txtCedula.Text & "'"
End If

vPaso = True
Call sbCargaGrid(vGrid, 7, strSQL)

vPaso = False

Me.MousePointer = vbDefault

Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarGridSeguimiento()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
            
vGridSeguimiento.MaxCols = 4
vGridSeguimiento.MaxRows = 0

If vContrato = Empty Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO" _
       & " from SIF_CONTROL_TAGS OT inner join SIF_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO" _
       & " where OT.codigo = '" & vPlan & "' and OT.documento = '" & vContrato & "' and OT.COD_Modulo = 'FND'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    vGridSeguimiento.MaxRows = vGridSeguimiento.MaxRows + 1
    vGridSeguimiento.Row = vGridSeguimiento.MaxRows
  
    vGridSeguimiento.Col = 1
    vGridSeguimiento.Text = rs!Descripcion
    vGridSeguimiento.TextTip = TextTipFixed
    vGridSeguimiento.TextTipDelay = 1000
    vGridSeguimiento.CellNote = "Usuario: " & rs!registro_usuario & "[" & rs!Registro_Fecha & "]"
            
    vGridSeguimiento.Col = 2
    vGridSeguimiento.Value = IIf(IsNull(rs!notas), "", rs!notas)
    
    vGridSeguimiento.Col = 3
    vGridSeguimiento.Value = IIf(IsNull(rs!Registro_Fecha), "", rs!Registro_Fecha)
    
    vGridSeguimiento.Col = 4
    vGridSeguimiento.Value = IIf(IsNull(rs!registro_usuario), "", rs!registro_usuario)
    
    vGridSeguimiento.RowHeight(vGridSeguimiento.Row) = vGridSeguimiento.MaxTextRowHeight(vGridSeguimiento.Row)
    rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargarListaErrores()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If txtCedula = Empty Then
    Exit Sub
End If

With lswErrores
 .ListItems.Clear
 
 strSQL = "select E.ID_ERROR,E.DESCRIPCION,ER.ID_ERROR as asignado, ISNULL(ER.APLICADO,'N') AS APLICADO, E.MENSAJE, ER.LINEA_ERR" _
        & " from sif_Omisiones E left join SIF_OMISIONESG ER on E.ID_ERROR = ER.ID_ERROR" _
        & " and ER.cedula = '" & txtCedula.Text & "' and ER.Modulo = 'FND' and ER.Codigo = '" & txtPlan.Text _
        & "' and ER.Documento = '" & txtContrato.Text & "'" _
        & " where E.ACTIVO = '1'  and E.ID_ERROR in(select ID_ERROR from SIF_OMISIONES_MODULOS where cod_modulo = 'FND') " _
        & " order by E.ID_ERROR"
        
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!ID_ERROR)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
         itmX.Tag = rs!LINEA_ERR
      End If
      itmX.SubItems(2) = rs!APLICADO
      itmX.SubItems(3) = rs!Mensaje
  rs.MoveNext
 Loop
 rs.Close
End With
End Sub


Private Sub sbCargarCombosEtiquetas()
Dim strSQL As String

On Error GoTo vError

    
    strSQL = "SELECT CT.TAG_CODIGO + ' - ' +  rtrim(CT.DESCRIPCION) as 'ItmX'" _
            & " FROM SIF_TAGS CT INNER JOIN SIF_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
            & " INNER JOIN SIF_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
            & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
            & "' and  CT.TAG_CODIGO in(select TAG_CODIGO from SIF_TAGS_MODULOS where cod_modulo = 'FND')" _
            & " order by CT.TAG_CODIGO"
    vPaso = True
    Call sbLlenaCbo(cboEtiquetas, strSQL, False, False)
    vPaso = False
    Call cboEtiquetas_Click
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaDetalle()
Dim strSQL As String, rs As New ADODB.Recordset

If vContrato = "" Then Exit Sub

On Error GoTo vError

strSQL = "select C.*,S.nombre,O.descripcion as Operadora,P.descripcion as PlanX" _
       & " from fnd_contratos C inner join Socios S on C.cedula = S.cedula" _
       & " inner join fnd_planes P on C.cod_plan = P.cod_plan and C.cod_operadora = P.cod_operadora" _
       & " inner join fnd_operadoras O on C.cod_operadora = O.cod_operadora" _
       & " where C.cod_operadora = " & vOperadora _
       & " and C.cod_plan = '" & vPlan & "' and C.cod_contrato = " & vContrato
Call OpenRecordSet(rs, strSQL)

 txtCedula.Text = rs!Cedula
 lblNombre.Caption = rs!Nombre
 txtOperadora = rs!Operadora
 txtOperadora.Tag = rs!Cod_Operadora
 txtDescripcion = rs!PlanX
 txtDescripcion.Tag = rs!cod_Plan
 lblContrato = rs!cod_contrato
 lblEstado = IIf((rs!Estado = "A"), "Activo", "Liquidado")
 lblFecha = Format(rs!FECHA_INICIO, "dd/mm/yyyy")

 lblMonto = Format(rs!Monto, "Standard")
 lblPlazo = rs!Plazo
 lblRenueva = IIf(rs!Renueva = "S", "SI", "NO")
 lblIncAnual = Format(rs!Inc_Anual, "Standard")
 lblIncTipo = IIf(rs!Inc_Tipo = "P", "Porcentaje", "Monto")
 lblAportes = Format(rs!aportes, "Standard")
 lblRendimiento = Format(rs!rendimiento, "Standard")
 
 lblOperacionASE = "Operación : " & IIf(IsNull(rs!Operacion), "", rs!Operacion)
 lblTotal.Caption = Format(CCur(lblAportes.Caption) + CCur(lblRendimiento.Caption), "Standard")
 rs.Close
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
 
End Sub

Private Sub tlbRefresh_ButtonClick(ByVal Button As MSComctlLib.Button)
Call TimerX_Timer
End Sub

Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)

If vPaso Then Exit Sub

On Error GoTo vError

If Col = 1 Then
    vGrid.Row = Row
    vGrid.Col = 5
    vContrato = vGrid.Text
    vGrid.Col = 6
    vPlan = vGrid.Text
    vGrid.Col = 7
    vOperadora = vGrid.Text
    
    txtPlan.Text = vPlan
    txtContrato.Text = vContrato
    
    SSTab1.Tab = 1
    Call SSTab1_Click(0)
End If

Exit Sub

vError:
vOperadora = 0
vPlan = ""
vContrato = 0
End Sub
