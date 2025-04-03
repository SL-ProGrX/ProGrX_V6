VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmCC_CA_Remesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos Automáticos: Remesas"
   ClientHeight    =   8220
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   12708
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12708
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   7965
      Width           =   12705
      _ExtentX        =   22416
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Casos:"
            TextSave        =   "Casos:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "Monto:"
            TextSave        =   "Monto:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "Autorizado:"
            TextSave        =   "Autorizado:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Remesa:"
            TextSave        =   "Remesa:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6615
      Left            =   15
      TabIndex        =   0
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22246
      _ExtentY        =   11663
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Envía"
      TabPicture(0)   =   "frmCC_CA_Remesas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpVence"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tlbEnvia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "vGrid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProceso"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboLinea"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboEntidad"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkTodosLosCasos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboCuotas"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Recibe / Aplica"
      TabPicture(1)   =   "frmCC_CA_Remesas.frx":07E1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(1)=   "Label2(1)"
      Tab(1).Control(2)=   "Label2(0)"
      Tab(1).Control(3)=   "tlbArchivo"
      Tab(1).Control(4)=   "vGridRecibe"
      Tab(1).Control(5)=   "tlbProceso"
      Tab(1).Control(6)=   "tlbX"
      Tab(1).Control(7)=   "txtArchivo"
      Tab(1).Control(8)=   "txtCasos"
      Tab(1).Control(9)=   "txtTotal"
      Tab(1).Control(10)=   "cboRemesa"
      Tab(1).Control(11)=   "chkExcel"
      Tab(1).Control(12)=   "optX(0)"
      Tab(1).Control(13)=   "optX(1)"
      Tab(1).Control(14)=   "optX(2)"
      Tab(1).Control(15)=   "fraAplica"
      Tab(1).ControlCount=   16
      Begin VB.ComboBox cboCuotas 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "frmCC_CA_Remesas.frx":0FB9
         Left            =   3000
         List            =   "frmCC_CA_Remesas.frx":0FBB
         Style           =   2  'Dropdown List
         TabIndex        =   32
         ToolTipText     =   "Cantidad de Cuotas Pendientes a Cobrar"
         Top             =   360
         Width           =   870
      End
      Begin VB.Frame fraAplica 
         Caption         =   "Aplicando Cargos Automáticos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   -71760
         TabIndex        =   26
         Top             =   1680
         Visible         =   0   'False
         Width           =   6255
         Begin MSComctlLib.ProgressBar ProgressBarX 
            Height          =   150
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   6015
            _ExtentX        =   10605
            _ExtentY        =   275
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblAplica 
            Alignment       =   1  'Right Justify
            Caption         =   "Detallando Abonos (Espere!)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   29
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label lblStatus 
            Caption         =   "...."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   5775
         End
      End
      Begin VB.OptionButton optX 
         Caption         =   "Aplicadas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   -71760
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optX 
         Caption         =   "Cerradas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   -73200
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optX 
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   -74760
         TabIndex        =   23
         Top             =   720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "Archivo Formato Excel?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66240
         TabIndex        =   21
         Top             =   6120
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkTodosLosCasos 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar solo Tarjetas Válidas?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   20
         Top             =   6120
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.ComboBox cboRemesa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   6135
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67440
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Monto"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtCasos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -65880
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboEntidad 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
      Begin VB.ComboBox cboLinea 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
      Begin VB.ComboBox cboProceso 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtArchivo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   5400
         Width           =   9615
      End
      Begin MSComctlLib.Toolbar tlbX 
         Height          =   264
         Left            =   -63840
         TabIndex        =   1
         Top             =   5400
         Width           =   1284
         _ExtentX        =   2265
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
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
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "info"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5052
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   12012
         _Version        =   524288
         _ExtentX        =   21188
         _ExtentY        =   8911
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
         MaxCols         =   491
         ScrollBars      =   2
         SpreadDesigner  =   "frmCC_CA_Remesas.frx":0FBD
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbProceso 
         Height          =   312
         Left            =   -65040
         TabIndex        =   13
         Top             =   360
         Width           =   2484
         _ExtentX        =   4382
         _ExtentY        =   550
         ButtonWidth     =   1778
         ButtonHeight    =   550
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
      Begin MSComctlLib.Toolbar tlbEnvia 
         Height          =   264
         Left            =   11280
         TabIndex        =   14
         Top             =   360
         Width           =   924
         _ExtentX        =   1630
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Información a Deducir"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cargar"
               Object.ToolTipText     =   "Cargar Deducciones"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridRecibe 
         Height          =   4212
         Left            =   -74880
         TabIndex        =   18
         Top             =   1080
         Width           =   12252
         _Version        =   524288
         _ExtentX        =   21611
         _ExtentY        =   7430
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
         MaxCols         =   491
         SpreadDesigner  =   "frmCC_CA_Remesas.frx":15CC
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbArchivo 
         Height          =   264
         Left            =   -68568
         TabIndex        =   22
         Top             =   360
         Width           =   924
         _ExtentX        =   1630
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Archivo"
               Object.ToolTipText     =   "Genera Archivo para el Banco"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cierra"
               Object.ToolTipText     =   "Cierra Remesa"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   315
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   1212
         _Version        =   1310720
         _ExtentX        =   2138
         _ExtentY        =   556
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
      Begin VB.Label Label2 
         Caption         =   "Cuotas:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   31
         ToolTipText     =   "Cantidad de Cuotas a Cobrar"
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Vence:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   30
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -67440
         TabIndex        =   17
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Remesa en Trámite:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   16
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Entidad:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7440
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Línea:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Proceso:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   -74640
         TabIndex        =   3
         Top             =   5400
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
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2100
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   9000
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   240
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_CA_Remesas.frx":1BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_CA_Remesas.frx":2610
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_CA_Remesas.frx":2DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_CA_Remesas.frx":37A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_CA_Remesas.frx":4168
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_CA_Remesas.frx":485F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_CA_Remesas.frx":51D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargos Automáticos (Tarjetas)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2280
      TabIndex        =   19
      Top             =   360
      Width           =   7092
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   10680
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmCC_CA_Remesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtTotal.Text = 0
    txtCasos.Text = 0
End Sub


Private Sub cboCliente_Change()
 Call sbLimpia
End Sub


Private Sub cboInstitucion_Change()
 Call sbLimpia
End Sub

Private Sub cboProceso_Change()
 Call sbLimpia
End Sub

Private Sub cboTipo_Change()
 Call sbLimpia
End Sub

Private Sub cboProceso_Click()
If vPaso Then Exit Sub
If cboProceso.ListCount = 0 Or cboProceso.Text = "" Then Exit Sub


Dim vFecha As Date

vFecha = Mid(cboProceso.Text, 1, 4) & "/" & Mid(cboProceso.Text, 5, 2) & "/01"
vFecha = DateAdd("d", -1, DateAdd("m", 1, vFecha))

If dtpVence.Value >= vFecha Then
   dtpVence.Value = DateAdd("d", -1, vFecha)
End If

dtpVence.MaxDate = vFecha
dtpVence.Value = vFecha

End Sub

Private Sub chkExcel_Click()
 Call sbLimpia
End Sub


Private Sub cboRemesa_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curTotal As Currency, iCasos As Long

If vPaso Or cboRemesa.ListCount = 0 Then Exit Sub

txtArchivo.Text = ""
txtTotal.Text = "0"
txtCasos.Text = "0"

curTotal = 0
iCasos = 0

strSQL = "exec spPrm_CA_Remesa_Consulta " & cboRemesa.ItemData(cboRemesa.ListIndex)
Call OpenRecordSet(rs, strSQL)
With vGridRecibe
 .MaxRows = 0
 Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  
  .Col = 1
  .Text = rs!Cedula
  .Col = 2
  .Text = rs!Nombre
  .Col = 3
  .Text = CStr(rs!Monto)
  .Col = 4
  .Text = rs!Tarjeta
  .Col = 5
  .Text = CStr(rs!Autorizacion & "")
  .Col = 6
  Select Case rs!Estado
    Case "P" 'Pendiente
      .Text = "Pendiente"
    Case "A" 'Autorizado
      .Text = "Autorizado"
    Case "X" 'Aplicado
      .Text = "Aplicado"
  End Select
  
  curTotal = curTotal + rs!Monto
  iCasos = .MaxRows
  
  rs.MoveNext
 Loop
End With
rs.Close

txtTotal.Text = Format(curTotal, "Standard")
txtCasos.Text = Format(iCasos, "###,###,##0")


End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer
Dim vProceso As Currency

vGrid.AppearanceStyle = fxGridStyle
vGridRecibe.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

strSQL = "select rtrim(COD_LINEA) + ' - ' + descripcion as 'ItmX' from PRM_CA_LINEAS where activo = 1"
Call sbLlenaCbo(cboLinea, strSQL, False, False)

strSQL = "select rtrim(cod_Entidad) + ' - ' + descripcion as 'ItmX' from PRM_CA_ENTIDAD where activo = 1"
Call sbLlenaCbo(cboEntidad, strSQL, False, False)

txtArchivo.Text = ""

vGrid.MaxCols = 5
vGrid.MaxRows = 0

vGridRecibe.MaxCols = 6
vGridRecibe.MaxRows = 0

vPaso = True

vProceso = GLOBALES.glngFechaCR
cboProceso.AddItem vProceso

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboProceso.AddItem vProceso
Next i
cboProceso.Text = GLOBALES.glngFechaCR

vPaso = False
Call cboProceso_Click


cboCuotas.AddItem "1"
cboCuotas.AddItem "2"
cboCuotas.AddItem "3"
cboCuotas.AddItem "4"
cboCuotas.AddItem "5"
cboCuotas.Text = "1"

ssTab.Tab = 0



End Sub

Private Sub sbCargaAutorizaciones()
Dim pTarjeta As String, pMonto As Currency, pComision As Currency, pAutorizacion As String
Dim pRemesa As Long, pFecha As Date, pReferencia As String
Dim strSQL As String, rsExcel As New ADODB.Recordset, strCadena As String
Dim fn

On Error GoTo vError

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboRemesa.ListCount <= 0 Then Exit Sub

If optX.Item(0).Value = False Then
   MsgBox "Solo se Pueden cargar Autorizaciones a una Remesa en estado Pendiente!", vbExclamation
   Exit Sub
End If

strSQL = ""
pRemesa = cboRemesa.ItemData(cboRemesa.ListIndex)
'TODO: Validar el Estado de la Remesa (Que no esté aplicada)

Me.MousePointer = vbHourglass

If chkExcel.Value = vbChecked Then

        
        Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
        
        With rsExcel
        
            Do While Not .EOF
              pTarjeta = Format(!Tarjeta, "########000000000000000")
              pMonto = !Monto
              pComision = !Comision
              pAutorizacion = !Autorizacion
              pFecha = !fecha
              pReferencia = !Referencia
              
              strSQL = strSQL & Space(10) & " exec spPrm_CA_Remesa_Autorizaciones " & pRemesa & ",'" & pTarjeta & "','" _
                     & pAutorizacion & "'," & pMonto & "," & pComision & ",'" & Format(pFecha, "yyyy/mm/dd") & "','" _
                     & pReferencia & "'"
              
                            
               If Len(strSQL) >= 20000 Then
                  Call ConectionExecute(strSQL)
                  strSQL = ""
               End If
               
                            
              .MoveNext
            Loop
            .Close
        End With
        


Else 'Archivo Texto
        fn = FreeFile
        Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
         Do While Not EOF(fn)
           Input #fn, strCadena
              'TODO: Falta Revisar el Formato del Banco
              pTarjeta = Mid(strCadena, 1, 14)
              pMonto = Mid(strCadena, 1, 14)
              pComision = Mid(strCadena, 1, 14)
              pAutorizacion = Mid(strCadena, 1, 14)
              pFecha = Mid(strCadena, 1, 14)
             
               If Len(strSQL) >= 20000 Then
                  Call ConectionExecute(strSQL)
                  strSQL = ""
               End If
         Loop
        Close #fn
        
End If 'Archivo Excel


'Procesa Lote Final
If Len(strSQL) >= 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


Me.MousePointer = vbDefault

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Call cboRemesa_Click

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call cboRemesa_Click

End Sub

Private Sub sbAplicar()
Dim strSQL As String, rs As New ADODB.Recordset, pRows As Long
Dim pTipoDoc As String, pNumDoc As String, pRemesa As Long, pProceso As Long, pRLinea As String
Dim vTotal As Long, vPendientes As Long, vProcesados As Long, vRegistros As Long


On Error GoTo vError

Me.MousePointer = vbHourglass

fraAplica.Visible = True

lblAplica.Caption = "Detallando Abonos (Espere!)"
lblAplica.Refresh

pRemesa = cboRemesa.ItemData(cboRemesa.ListIndex)

'Borra El Historial en el Proceso de Creditos
strSQL = "exec  spPrm_CA_Remesa_Aplica_Inicializa " & pRemesa & ",'" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If glogon.error Then
   rs.Close
   Me.MousePointer = vbDefault
   Exit Sub
End If
    
    pTipoDoc = rs!TipoDoc
    pNumDoc = rs!NumDoc
    pProceso = rs!Proceso
    pRLinea = rs!RLinea

rs.Close

lblStatus.Caption = "Detallando los Abonos a créditos..."
lblStatus.Refresh


strSQL = "exec spPrm_CA_Abonos_Detalla_Main " & pRemesa & ",'" & pRLinea & "'," & pProceso & ",1,50"
Call OpenRecordSet(rs, strSQL)
  vTotal = rs!total
  vPendientes = IIf(IsNull(rs!Pendientes) = True, 0, rs!Pendientes)
  vProcesados = IIf(IsNull(rs!Procesados) = True, 0, rs!Procesados)
rs.Close

ProgressBarX.Max = vTotal + 1
ProgressBarX.Value = vProcesados

  
lblStatus.Caption = "Detallando..Registro # " & ProgressBarX.Value & " de " & ProgressBarX.Max & "     " & Format((ProgressBarX.Value / ProgressBarX.Max) * 100, "##0") & "%"
lblStatus.Refresh

Do While vPendientes > 0
    strSQL = "exec spPrm_CA_Abonos_Detalla_Main " & pRemesa & ",'" & pRLinea & "'," & pProceso & ",0,50"
    Call OpenRecordSet(rs, strSQL)
    
        vTotal = rs!total
        vPendientes = rs!Pendientes
        vProcesados = rs!Procesados
    
    rs.Close

  ProgressBarX.Value = vProcesados
  lblStatus.Caption = "Cargando..Registro # " & ProgressBarX.Value & " de " & ProgressBarX.Max & "     " & Format((ProgressBarX.Value / ProgressBarX.Max) * 100, "##0") & "%"
  lblStatus.Refresh

Loop


lblAplica.Caption = "Aplicado Abonos (Espere!)"
lblAplica.Refresh

lblStatus.Caption = "Aplicando a Abonos..."
lblStatus.Refresh


strSQL = "select count(*) + 1 as Total from prm_ca_creditos" _
       & " where cod_Remesa = " & pRemesa _
       & " and id_aplicacion = 1 and ind_paso = 0"
Call OpenRecordSet(rs, strSQL)
    ProgressBarX.Value = 0
    ProgressBarX.Max = rs!total
    vRegistros = rs!total
rs.Close

Do While vRegistros > 0
   strSQL = "exec spPrm_CA_Abonos_Aplica " & pRemesa & "," & pProceso & ",'" & glogon.Usuario & "','" & pTipoDoc & "','" & pNumDoc & "',50"
   Call OpenRecordSet(rs, strSQL)
     vRegistros = rs.Fields(0).Value
   rs.Close
   
  If (ProgressBarX.Value + vRegistros) <= ProgressBarX.Max Then ProgressBarX.Value = ProgressBarX.Value + vRegistros
  lblStatus.Caption = "Aplicando Registro # " & ProgressBarX.Value & " de " & ProgressBarX.Max & "     " & Format((ProgressBarX.Value / ProgressBarX.Max) * 100, "##0") & "%"
  lblStatus.Refresh
Loop


'Aplica Inconsistencia a Fondos
strSQL = "exec spPrm_CA_Aplica_Fondos_Main " & pRemesa & ",'" & glogon.Usuario & "','" & pTipoDoc & "','" & pNumDoc & "'"
Call ConectionExecute(strSQL)

'Cierra Remesa
strSQL = "update prm_Ca_Remesas set Estado = 'A'  where COD_REMESA = " & pRemesa
Call ConectionExecute(strSQL)

'Asiento
strSQL = "exec spPrm_CA_Aplica_Asiento '" & pTipoDoc & "','" & pNumDoc & "','" & glogon.Usuario & "'," & pRemesa
Call ConectionExecute(strSQL)

fraAplica.Visible = False

Me.MousePointer = vbDefault

MsgBox "Remesa fue aplicada satisfactoriamente con " & pTipoDoc & " No: " & pNumDoc, vbInformation

Call sbImprimeRecibo(pNumDoc, pTipoDoc)

Call optX_Click(0)

Exit Sub


vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub optX_Click(Index As Integer)

Call ssTab_Click(1)

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

If ssTab.Tab = 1 Then
  vGridRecibe.MaxRows = 0
  
  txtTotal.Text = 0
  txtCasos.Text = 0
  
  vPaso = True
      strSQL = "exec spPrm_CA_Remesa_Lista "
      
      Select Case True
         Case optX.Item(0).Value 'Pendientes
                strSQL = strSQL & "'P'"
         Case optX.Item(1).Value 'Cerradas
                strSQL = strSQL & "'C'"
         Case optX.Item(2).Value 'Aplicadas
                strSQL = strSQL & "'A'"
      End Select
      
      Call sbLlenaCbo(cboRemesa, strSQL, False, True)
  vPaso = False
  
  Call cboRemesa_Click
End If


End Sub


Private Sub sbArchivo_Banco(pRemesa As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, vTempo As String, vArchivo As String, vRuta As String, vFile As String
Dim fnFile

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla


fnFile = FreeFile


Me.MousePointer = vbHourglass


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Cargo Automático\"


On Error GoTo vError

vArchivo = Replace(cboRemesa.Text, ":", "_")
vArchivo = Replace(vArchivo, ".", "-") & ".txt"
vRuta = SIFGlobal.DirectorioDeResultados & "\Cargo Automático"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


strSQL = "exec spPrm_CA_Remesa_Archivo_Envia " & pRemesa
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 
 Select Case rs!Formato
     Case "BAC"
                vCadena = "1,"
                vCadena = vCadena & Mid(rs!Cedula, 1, 25) & ","
                vCadena = vCadena & Mid(rs!Referencia, 1, 25) & ","
                vCadena = vCadena & Trim(rs!Tarjeta) & ","
                vCadena = vCadena & Trim(rs!Tarjeta_Vence) & ","
                vCadena = vCadena & Format(rs!Monto, "########0.##") & ","
                vCadena = vCadena & Format(rs!Fecha_Vence, "mmyyyy") & ","
                vCadena = vCadena & Format(rs!Fecha_Transaccion, "ddmmyyyy") & ","
                vCadena = vCadena & Mid(Trim(rs!Email), 1, 30) & ","
                vCadena = vCadena & Mid(Trim(rs!Nombre), 1, 30)
    
     Case "BNCR"
                vCadena = SIFGlobal.fxStringRelleno(rs!Tarjeta, "I", "0", 16)
                vCadena = vCadena & Format(rs!Monto, "0000000.00")
                vCadena = vCadena & SIFGlobal.fxStringRelleno(rs!NUMERO_AFILIADO, "I", "0", 10)
                vCadena = vCadena & Format(rs!Fecha_Transaccion, "ddmmyyyy")
                vCadena = vCadena & Format(rs!Fecha_Vence, "mmyyyy")
                vCadena = vCadena & SIFGlobal.fxStringRelleno(rs!Referencia, "I", " ", 40)
     Case Else
                vCadena = SIFGlobal.fxStringRelleno(rs!Tarjeta, "I", "0", 16)
                vCadena = vCadena & Format(rs!Monto, "0000000.00")
                vCadena = vCadena & SIFGlobal.fxStringRelleno(rs!NUMERO_AFILIADO, "I", "0", 10)
                vCadena = vCadena & Format(rs!Fecha_Transaccion, "ddmmyyyy")
                vCadena = vCadena & Format(rs!Fecha_Vence, "mmyyyy")
                vCadena = vCadena & SIFGlobal.fxStringRelleno(rs!Referencia, "I", " ", 40)
 End Select
 
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



Private Sub sbCierra(pRemesa As Long)
Dim strSQL As String, pRow As Long, i As Integer

On Error GoTo vError

 i = MsgBox("Esta seguro que desea cerrar la Remesa?", vbYesNo)
 If i = vbNo Then Exit Sub
 
strSQL = "exec spPrm_CA_Remesa_Cierra " & pRemesa & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL, , pRow)

If pRow <> -1 Then
    MsgBox "La Remesa " & pRemesa & " ha sido cerrada con éxito", vbInformation
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbArchivo_ButtonClick(ByVal Button As MSComctlLib.Button)

If cboRemesa.ListCount = 0 Then Exit Sub

Select Case Button.Key
  Case "Archivo"
        Call sbArchivo_Banco(cboRemesa.ItemData(cboRemesa.ListIndex))
  Case "Cierra"
        Call sbCierra(cboRemesa.ItemData(cboRemesa.ListIndex))
End Select
End Sub

Private Sub tlbEnvia_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCedula As String, pNombre As String, pCompromiso As Currency, pTarjeta As String, pTarjetaVence As Date
Dim x As Long, pRemesa As Long, vFechaInicio As Date

On Error GoTo vError

Select Case Button.Key
   Case "Buscar"
        strSQL = "exec spPrm_CA_Busca_Casos " & cboProceso.Text & ",'" & SIFGlobal.fxCodText(cboLinea.Text) _
               & "'," & chkTodosLosCasos.Value & ",'" & Format(dtpVence.Value, "yyyy/mm/dd") & "'," & cboCuotas.Text
        Call sbCargaGrid(vGrid, 5, strSQL, True)
        
        vGrid.MaxRows = vGrid.MaxRows - 1
   
   Case "Cargar"
  


    vFechaInicio = Mid(cboProceso.Text, 1, 4) & "/" & Mid(cboProceso.Text, 5, 2) & "/01"
   
   strSQL = "exec spPrm_CA_Remesa " & cboProceso.Text & ",'" & SIFGlobal.fxCodText(cboLinea.Text) & "','" & SIFGlobal.fxCodText(cboEntidad.Text) _
           & "','" & glogon.Usuario & "','" & Format(vFechaInicio, "yyyy/mm/dd") & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "'," & cboCuotas.Text
   Call OpenRecordSet(rs, strSQL)
        pRemesa = rs!Remesa
   rs.Close
   
   Me.MousePointer = vbHourglass
   
   strSQL = ""
   With vGrid
      For x = 1 To .MaxRows
            .Row = x
            .Col = 1
            pCedula = .Text
            .Col = 2
            pNombre = .Text
            .Col = 3
            pCompromiso = CCur(.Text)
            .Col = 4
            pTarjeta = .Text
            .Col = 5
            pTarjetaVence = IIf((.Text = ""), Date, CDate(.Text))
            
            strSQL = strSQL & Space(10) & "exec spPrm_CA_Remesa_Detalle " & pRemesa & ",'" & Trim(pCedula) & "','" & pNombre _
                    & "'," & pCompromiso & ",'" & pTarjeta & "','" & Format(pTarjetaVence, "yyyy/mm/dd") & "'"
            
            If Len(strSQL) > 20000 Then
              Call ConectionExecute(strSQL)
              strSQL = ""
            End If
   
      Next x
   End With
   
    'Lote Pendiente
    If Len(strSQL) > 0 Then
      Call ConectionExecute(strSQL)
      strSQL = ""
    End If
    
    Me.MousePointer = vbDefault
    MsgBox "Información para Cobro Automático realizada satisfactoriamente!", vbInformation
    
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "aplicar"
    If vGridRecibe.MaxRows = 0 Then
       MsgBox "No existe Detalle a Procesar Remesa...[verifique!]", vbExclamation
       Exit Sub
    End If
    
    If optX.Item(1).Value = False Then
       MsgBox "Solo se pueden aplicar Remesas en Estado: Cerradas...[verifique!]", vbExclamation
       Exit Sub
    End If
    
    Call sbAplicar
  
  Case "cancelar"
   Call cboRemesa_Click

End Select

End Sub


Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "buscar"
        
        txtArchivo.Text = ""
        
        With Cmd
         If chkExcel.Value = vbChecked Then
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
                .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
                .ShowOpen
        
                If .FileName = "" Then
                    MsgBox "Archivo no válido...", vbExclamation
                    Exit Sub
                End If
        
                If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
                    'Ok
                Else
                    MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                    Exit Sub
                End If
                
         Else
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo del Banco [Texto]..."
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
         End If
        
         txtArchivo.Text = .FileName
        
        End With

  Case "cargar"
    Call sbCargaAutorizaciones
  
  
  Case "info"
      Dim vMensaje As String
      
      vMensaje = "Formato del Archivo: Microsoft EXCEL 97-2003" & vbCrLf _
                & vbCrLf & " - Nombre de la Hoja: Import" _
                & vbCrLf & " - Nombres de los Campos en la Primer Fila" _
                & vbCrLf & " - Campos: REFERENCIA, TARJETA, MONTO, AUTORIZACION, FECHA, COMISION"
     
      MsgBox vMensaje, vbInformation
     
End Select

End Sub


