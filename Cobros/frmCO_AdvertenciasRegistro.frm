VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCO_AdvertenciasRegistro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro / Resolución de Advertencias"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9528
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9528
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   7080
      Top             =   120
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4092
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   9252
      _Version        =   1245187
      _ExtentX        =   16319
      _ExtentY        =   7218
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Historial"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   15
      Item(1).Control(0)=   "Label2(0)"
      Item(1).Control(1)=   "Label2(1)"
      Item(1).Control(2)=   "Label2(2)"
      Item(1).Control(3)=   "Label2(3)"
      Item(1).Control(4)=   "FlatScrollBar"
      Item(1).Control(5)=   "txtAdvCod"
      Item(1).Control(6)=   "txtAdvDesc"
      Item(1).Control(7)=   "txtNotas"
      Item(1).Control(8)=   "txtLinea"
      Item(1).Control(9)=   "txtEstado"
      Item(1).Control(10)=   "dtpVence"
      Item(1).Control(11)=   "Label2(8)"
      Item(1).Control(12)=   "btnRegistro(0)"
      Item(1).Control(13)=   "btnRegistro(1)"
      Item(1).Control(14)=   "btnRegistro(2)"
      Item(2).Caption =   "Resolución"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "Label2(4)"
      Item(2).Control(1)=   "Label2(5)"
      Item(2).Control(2)=   "Label2(6)"
      Item(2).Control(3)=   "Label2(7)"
      Item(2).Control(4)=   "txtResNotas"
      Item(2).Control(5)=   "txtResLinea"
      Item(2).Control(6)=   "cboResEstado"
      Item(2).Control(7)=   "txtResAdvCod"
      Item(2).Control(8)=   "txtResAdvertencia"
      Item(2).Control(9)=   "btnAplicar"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3612
         Left            =   -70000
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   9252
         _Version        =   1245187
         _ExtentX        =   16319
         _ExtentY        =   6371
         _StockProps     =   77
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnRegistro 
         Height          =   372
         Index           =   0
         Left            =   5160
         TabIndex        =   28
         Top             =   3600
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Nuevo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCO_AdvertenciasRegistro.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtResAdvertencia 
         Height          =   312
         Left            =   -66880
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1245187
         _ExtentX        =   9758
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtResAdvCod 
         Height          =   312
         Left            =   -68200
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAdvDesc 
         Height          =   312
         Left            =   3120
         TabIndex        =   12
         Top             =   1080
         Width           =   5532
         _Version        =   1245187
         _ExtentX        =   9758
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAdvCod 
         Height          =   312
         Left            =   1800
         TabIndex        =   11
         Top             =   1080
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1392
         Left            =   1800
         TabIndex        =   13
         Top             =   1440
         Width           =   6852
         _Version        =   1245187
         _ExtentX        =   12086
         _ExtentY        =   2455
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLinea 
         Height          =   312
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   252
         Left            =   8760
         TabIndex        =   15
         Top             =   1080
         Width           =   492
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtResNotas 
         Height          =   1392
         Left            =   -68200
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   6852
         _Version        =   1245187
         _ExtentX        =   12086
         _ExtentY        =   2455
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtResLinea 
         Height          =   312
         Left            =   -68200
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboResEstado 
         Height          =   312
         Left            =   -63280
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   6720
         TabIndex        =   25
         Top             =   600
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   312
         Left            =   7320
         TabIndex        =   27
         Top             =   3000
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnRegistro 
         Height          =   372
         Index           =   1
         Left            =   6240
         TabIndex        =   29
         Top             =   3600
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Borrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCO_AdvertenciasRegistro.frx":0632
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRegistro 
         Height          =   372
         Index           =   2
         Left            =   7560
         TabIndex        =   30
         Top             =   3600
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCO_AdvertenciasRegistro.frx":0BD6
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   372
         Left            =   -62680
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aplicar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCO_AdvertenciasRegistro.frx":1307
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   8
         Left            =   3600
         TabIndex        =   26
         Top             =   3000
         Width           =   3612
         _Version        =   1245187
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Vencimiento de la advertencia:  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   7
         Left            =   -64000
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   6
         Left            =   -69520
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   5
         Left            =   -69520
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Advertencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   4
         Left            =   -69520
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Línea Id"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   3
         Left            =   6000
         TabIndex        =   10
         Top             =   600
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Advertencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Línea Id"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   5496
      Width           =   9528
      _ExtentX        =   16806
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3598
            MinWidth        =   3598
            Object.ToolTipText     =   "Usuario de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha Resolución"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Usuario Resolución"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   5652
      _Version        =   1245187
      _ExtentX        =   9970
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   7
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   1812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1812
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_AdvertenciasRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean, vFecha As Date

Private Sub btnAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMovimiento As String

On Error GoTo vError

If Mid(txtEstado.Text, 1, 1) <> "A" Then
   MsgBox "La advertencia no se puede alterar porque ya fue resuelta!", vbExclamation
   Exit Sub
End If


If txtResAdvCod.Text = "" Or Len(txtResNotas.Text) <= 10 Or txtResLinea.Text = "" Then
  MsgBox "No se ha especificado ninguna advertencia o las notas no son válidas!", vbExclamation
  Exit Sub
End If
        
        
strSQL = "update CBR_ADVERTENCIAS_CASOS set Estado = '" & Mid(cboResEstado.Text, 1, 1) & "', Resolucion_Fecha = dbo.MyGetdate()" _
       & ", Resolucion_Usuario = '" & glogon.Usuario & "', Resolucion_Notas = '" & txtResNotas.Text _
       & "' where cod_Advertencia = '" & txtResAdvCod.Text & "' and Linea = " & txtResLinea.Text _
       & " and cedula = '" & txtCedula.Text & "'"
Call ConectionExecute(strSQL)
        
Call Bitacora("Aplica", "Advertencia ..: Id.(" & txtResLinea.Text & ") Cod.(" & txtResAdvCod.Text & ") Ced." & txtCedula.Text & " Est.:" & cboResEstado.Text)

MsgBox "Advertencia Resuelta satisfactorianmente...!", vbInformation

Call sbConsultaAdv(txtResAdvCod.Text, txtResLinea.Text)
        
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRegistro_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMovimiento As String

On Error GoTo vError

If Mid(txtEstado.Text, 1, 1) <> "A" Then
   MsgBox "La advertencia no se puede alterar porque ya fue resuelta!", vbExclamation
   Exit Sub
End If

Call sbSIFCleanTxtInject(txtNotas)


Select Case Index
    Case 0 'nuevo
        Call sbLimpia
        
    Case 1 'borrar
    
        If txtAdvCod.Text = "" Or txtLinea.Text = "" Then
          MsgBox "No se ha especificado ninguna advertencia!", vbExclamation
          Exit Sub
        End If
        
        strSQL = "delete CBR_ADVERTENCIAS_CASOS where Cedula = '" & txtCedula.Text & "' and cod_Advertencia = '" _
               & txtAdvCod.Text & "' and Linea = " & txtLinea.Text
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Advertencia ..: Id.(" & txtLinea.Text & ") Cod.(" & txtAdvCod.Text & ") Ced." & txtCedula.Text)
        
        MsgBox "Advertencia Eliminada...!", vbInformation
        
        Call sbConsulta
    
    Case 2 'guardar
    
        If txtAdvCod.Text = "" Or Len(txtNotas.Text) <= 10 Then
          MsgBox "No se ha especificado ninguna advertencia o las notas no son válidas!", vbExclamation
          Exit Sub
        End If
        
        If txtLinea.Text = "" Then
           vMovimiento = "Inserta"
           strSQL = "exec spCbrAdvertenciaRegistro '" & txtCedula.Text & "','" & txtAdvCod.Text & "','" & Format(dtpVence.Value, "yyyy/mm/dd") _
                  & "','" & glogon.Usuario & "','" & txtNotas.Text & "',0"
        Else
           vMovimiento = "Modifica"
           strSQL = "exec spCbrAdvertenciaRegistro '" & txtCedula.Text & "','" & txtAdvCod.Text & "','" & Format(dtpVence.Value, "yyyy/mm/dd") _
                  & "','" & glogon.Usuario & "','" & txtNotas.Text & "'," & txtLinea.Text
        
        End If
        
        Call OpenRecordSet(rs, strSQL)
           txtLinea.Text = rs!Linea
        rs.Close
        
        Call Bitacora(vMovimiento, "Advertencia ..: Id.(" & txtLinea.Text & ") Cod.(" & txtAdvCod.Text & ") Ced." & txtCedula.Text)
        
        MsgBox "Advertencia Registrada/Actualizada satisfactorianmente...!", vbInformation
        
        Call sbConsultaAdv(txtAdvCod.Text, txtLinea.Text)
        
End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_Advertencia,descripcion from CBR_ADVERTENCIAS_TIPO"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_Advertencia > '" & txtAdvCod.Text & "' and Activa = 1" _
              & " order by cod_Advertencia asc"
    Else
       strSQL = strSQL & " where cod_Advertencia < '" & txtAdvCod.Text & "' and Activa = 1" _
              & " order by cod_Advertencia desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtAdvCod.Text = rs!COD_ADVERTENCIA
      txtAdvDesc.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 4

End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

txtCedula.Text = GLOBALES.gTag

cboResEstado.Clear
cboResEstado.AddItem "Activa"
cboResEstado.AddItem "Resuelta"
cboResEstado.AddItem "Descartada"
cboResEstado.Text = "Activa"


With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 900
    .Add , , "Advertencia", 2500
    .Add , , "Estado", 1300, vbCenter
    .Add , , "Reg. Fecha", 1800, vbCenter
    .Add , , "Reg. Usuario", 1800, vbCenter
    .Add , , "Reg. Notas", 2500
    .Add , , "Res. Fecha", 1800, vbCenter
    .Add , , "Res. Usuario", 1800, vbCenter
    .Add , , "Res. Notas", 2500
End With


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub
 
Call sbConsultaAdv(Item.Tag, Item.Text)

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 
If Item.Index = 0 Then
 Call sbConsulta
End If

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

tcMain.Item(0).Selected = True
vFecha = DateAdd("d", 7, fxFechaServidor)

If txtCedula.Text <> "" Then
    Call sbConsulta(1)
Else
    txtCedula.SetFocus
End If

End Sub

Private Sub txtAdvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAdvDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_Advertencia"
  gBusquedas.Orden = "cod_Advertencia"
  gBusquedas.Consulta = "select cod_Advertencia,Descripcion from CBR_ADVERTENCIAS_TIPO"
  gBusquedas.Filtro = " and Activa =  1"
  frmBusquedas.Show vbModal
  txtAdvCod.Text = gBusquedas.Resultado
  txtAdvDesc.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtAdvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select cod_Advertencia,Descripcion from CBR_ADVERTENCIAS_TIPO"
  gBusquedas.Filtro = " and Activa =  1"
  frmBusquedas.Show vbModal
  txtAdvCod.Text = gBusquedas.Resultado
  txtAdvDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub sbLimpia()

txtLinea.Text = ""
txtAdvCod.Text = ""
txtAdvCod.Tag = ""
txtAdvDesc.Text = ""
txtNotas.Text = ""
dtpVence.Value = vFecha

txtEstado.Text = "Activa"

txtResLinea.Text = ""
cboResEstado.Text = "Resuelta"
txtResAdvertencia.Text = ""
txtResAdvCod.Text = ""
txtResAdvCod.Tag = ""
txtResNotas.Text = ""

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(3).Text = ""
StatusBarX.Panels(4).Text = ""
        
End Sub



Private Sub sbConsulta(Optional pInicial As Integer = 0)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

vPaso = True

Call sbLimpia

If pInicial = 1 Then
   txtNombre.Text = fxNombre(txtCedula.Text)
End If



tcMain.Item(0).Selected = True

lsw.ListItems.Clear

'Advertencias Registradas
strSQL = "select Cs.*,Tp.Descripcion as 'AdvertenciaDesc'" _
       & " from CBR_ADVERTENCIAS_CASOS Cs inner join CBR_ADVERTENCIAS_TIPO Tp on Cs.cod_Advertencia = Tp.cod_Advertencia" _
       & " where Cs.Cedula = '" & txtCedula.Text & "' order by Cs.Estado,Cs.Registro_Fecha desc"
Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Linea)
      itmX.SubItems(1) = rs!AdvertenciaDesc
      
      Select Case rs!Estado
        Case "A" 'Activa
          itmX.SubItems(2) = "Activa"
        Case "R" 'Resuelta
          itmX.SubItems(2) = "Resuelta"
        Case "D" 'Descartada
          itmX.SubItems(2) = "Descartada"
      End Select
      itmX.SubItems(3) = rs!Registro_Fecha
      itmX.SubItems(4) = rs!Registro_Usuario
      itmX.SubItems(5) = rs!notas
      itmX.SubItems(6) = rs!Resolucion_Fecha & ""
      itmX.SubItems(7) = rs!Resolucion_Usuario & ""
      itmX.SubItems(8) = rs!Resolucion_Notas & ""
      
      itmX.Tag = rs!COD_ADVERTENCIA
  rs.MoveNext
Loop
rs.Close

vPaso = False
Me.MousePointer = vbDefault

End Sub



Private Sub sbConsultaAdv(pAdvCod As String, pAdvLinea As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem


Me.MousePointer = vbHourglass

Call sbLimpia

'Advertencias Registradas
strSQL = "select Cs.*,Tp.Descripcion as 'AdvertenciaDesc'" _
       & " from CBR_ADVERTENCIAS_CASOS Cs inner join CBR_ADVERTENCIAS_TIPO Tp on Cs.cod_Advertencia = Tp.cod_Advertencia" _
       & " where Cs.Cedula = '" & txtCedula.Text & "' and Cs.cod_Advertencia = '" & pAdvCod & "' and Linea = " & pAdvLinea
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then


    txtLinea.Text = rs!Linea
    txtAdvCod.Text = rs!COD_ADVERTENCIA
    txtAdvDesc.Text = rs!AdvertenciaDesc
    txtNotas.Text = rs!notas
    dtpVence.Value = rs!Fecha_Vence
    
    txtResLinea.Text = rs!Linea
    txtResAdvertencia.Text = rs!AdvertenciaDesc
    txtResAdvCod.Text = rs!COD_ADVERTENCIA
    txtResNotas.Text = rs!Resolucion_Notas & ""
        
    StatusBarX.Panels(1).Text = rs!Registro_Fecha
    StatusBarX.Panels(2).Text = rs!Registro_Usuario
    StatusBarX.Panels(3).Text = rs!Resolucion_Fecha & ""
    StatusBarX.Panels(4).Text = rs!Resolucion_Usuario & ""

      Select Case rs!Estado
        Case "A" 'Activa
          txtEstado.Text = "Activa"
          tcMain.Item(1).Selected = True
          
        Case "R" 'Resuelta
          txtEstado.Text = "Resuelta"
          cboResEstado.Text = "Resuelta"
          tcMain.Item(2).Selected = True
        Case "D" 'Descartada
          txtEstado.Text = "Descartada"
          cboResEstado.Text = "Descartada"
          tcMain.Item(2).Selected = True
      End Select

End If
rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbConsulta(1)
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    
    frmBusquedas.Show vbModal
    txtCedula = Trim(gBusquedas.Resultado)
    gBusquedas.Consulta = ""
    gBusquedas.Columna = ""
    gBusquedas.Orden = ""
    gBusquedas.Resultado = ""
    
    
    If Trim(txtCedula) <> "" Then
        Call sbConsulta(1)
    End If

End If


Exit Sub

vError:

End Sub
