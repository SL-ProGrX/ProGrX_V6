VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_Telemarketing_Consultas 
   Caption         =   "Mercadeo Digital: Consultas"
   ClientHeight    =   8010
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox gBox_Procesando 
      Height          =   1332
      Left            =   5640
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   5892
      _Version        =   1441793
      _ExtentX        =   10393
      _ExtentY        =   2350
      _StockProps     =   79
      Caption         =   "Procesando: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   252
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   3972
         _Version        =   1441793
         _ExtentX        =   7011
         _ExtentY        =   450
         _StockProps     =   93
         Scrolling       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         MarqueeDelay    =   60
      End
      Begin VB.Label lblProcesando 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bajando Información..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   3972
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl_Main 
      Height          =   7812
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12612
      _Version        =   1441793
      _ExtentX        =   22246
      _ExtentY        =   13779
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
      PaintManager.Position=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "Colocación"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "GroupBox1"
      Item(0).Control(1)=   "vGrid"
      Item(1).Caption =   "Clientes en Común"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "GroupBox2"
      Item(1).Control(1)=   "vGrid_Clientes"
      Item(1).Control(2)=   "vGrid_Operaciones"
      Item(2).Caption =   "Contactos"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "vGrid_Contactos"
      Item(2).Control(1)=   "GroupBox3"
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   6972
         Left            =   360
         TabIndex        =   47
         Top             =   120
         Width           =   4212
         _Version        =   1441793
         _ExtentX        =   7429
         _ExtentY        =   12298
         _StockProps     =   79
         Caption         =   "Filtro para Personas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   252
            Left            =   3960
            TabIndex        =   67
            Top             =   5280
            Width           =   252
            _Version        =   1441793
            _ExtentX        =   444
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbMsj 
            Height          =   372
            Index           =   0
            Left            =   1320
            TabIndex        =   62
            Top             =   4080
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "SMS"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker dtpPersonaInicio 
            Height          =   312
            Left            =   1440
            TabIndex        =   48
            Top             =   1200
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   550
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtpPersonaCorte 
            Height          =   312
            Left            =   1440
            TabIndex        =   49
            Top             =   1560
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   550
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.PushButton btnPersonas 
            Height          =   732
            Left            =   1680
            TabIndex        =   54
            Top             =   6120
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   1291
            _StockProps     =   79
            Caption         =   "Buscar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_Telemarketing_Consultas.frx":0000
            TextImageRelation=   1
         End
         Begin XtremeSuiteControls.PushButton btnPersonas_Exportar 
            Height          =   732
            Left            =   2760
            TabIndex        =   55
            Top             =   6120
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   1291
            _StockProps     =   79
            Caption         =   "Exportar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_Telemarketing_Consultas.frx":0A1E
            TextImageRelation=   1
         End
         Begin XtremeSuiteControls.ComboBox cboPersonaEstado 
            Height          =   312
            Left            =   1440
            TabIndex        =   56
            Top             =   2160
            Width           =   2532
            _Version        =   1441793
            _ExtentX        =   4471
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboPersonaFecha 
            Height          =   312
            Left            =   1440
            TabIndex        =   57
            Top             =   840
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   372
            Left            =   1440
            TabIndex        =   58
            Top             =   2520
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Sin Morosidad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox CheckBox2 
            Height          =   372
            Left            =   1440
            TabIndex        =   59
            Top             =   2880
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Morosos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox CheckBox3 
            Height          =   372
            Left            =   1440
            TabIndex        =   60
            Top             =   3240
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Traslado de Deudas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox CheckBox4 
            Height          =   372
            Left            =   1440
            TabIndex        =   61
            Top             =   3600
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cobro Judicial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Value           =   1
         End
         Begin XtremeSuiteControls.RadioButton rbMsj 
            Height          =   372
            Index           =   1
            Left            =   1320
            TabIndex        =   63
            Top             =   4440
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Whatsapp"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton rbMsj 
            Height          =   372
            Index           =   2
            Left            =   1320
            TabIndex        =   64
            Top             =   4800
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Email"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   696
            Left            =   1320
            TabIndex        =   66
            Top             =   5280
            Width           =   2532
            _Version        =   1441793
            _ExtentX        =   4466
            _ExtentY        =   1228
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mensaje:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   12
            Left            =   360
            TabIndex        =   65
            Top             =   5280
            Width           =   1212
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Corte"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   11
            Left            =   480
            TabIndex        =   53
            Top             =   1560
            Width           =   852
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   10
            Left            =   480
            TabIndex        =   52
            Top             =   1200
            Width           =   852
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   132
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   2532
            _Version        =   1441793
            _ExtentX        =   4466
            _ExtentY        =   233
            _StockProps     =   79
            Caption         =   "Fecha Referencia:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   7
            Left            =   360
            TabIndex        =   50
            Top             =   2160
            Width           =   1212
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   7212
         Left            =   -69640
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   4572
         _Version        =   1441793
         _ExtentX        =   8064
         _ExtentY        =   12721
         _StockProps     =   79
         Caption         =   "Seleccione las líneas a combinar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkIntegral 
            Height          =   612
            Left            =   240
            TabIndex        =   41
            ToolTipText     =   "Agrupa los códigos de Créditos y Recaudación para coincidencia en ambos grupos"
            Top             =   6480
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Analisis Grupal: Créditos vrs Recaudación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Alignment       =   1
         End
         Begin VB.TextBox txtCrd_Consulta 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   4212
         End
         Begin MSComctlLib.ListView lswCrd 
            Height          =   2172
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   4212
            _ExtentX        =   7435
            _ExtentY        =   3836
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   1658
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   7832
            EndProperty
         End
         Begin MSComctlLib.ListView lswSel 
            Height          =   2052
            Left            =   120
            TabIndex        =   33
            Top             =   4320
            Width           =   4212
            _ExtentX        =   7435
            _ExtentY        =   3625
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   1658
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   7832
            EndProperty
         End
         Begin XtremeSuiteControls.RadioButton OptLineas 
            Height          =   252
            Index           =   0
            Left            =   480
            TabIndex        =   34
            Top             =   480
            Width           =   2652
            _Version        =   1441793
            _ExtentX        =   4678
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Líneas de Crédito"
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
         Begin XtremeSuiteControls.RadioButton OptLineas 
            Height          =   252
            Index           =   1
            Left            =   480
            TabIndex        =   35
            Top             =   840
            Width           =   2652
            _Version        =   1441793
            _ExtentX        =   4678
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Códigos de Recaudación"
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
         Begin XtremeSuiteControls.PushButton btnClienteBuscar 
            Height          =   732
            Left            =   2160
            TabIndex        =   37
            Top             =   6480
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   1291
            _StockProps     =   79
            Caption         =   "Buscar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_Telemarketing_Consultas.frx":1223
            TextImageRelation=   1
         End
         Begin XtremeSuiteControls.PushButton btnClienteExportar 
            Height          =   732
            Left            =   3240
            TabIndex        =   38
            Top             =   6480
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   1291
            _StockProps     =   79
            Caption         =   "Exportar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_Telemarketing_Consultas.frx":1C41
            TextImageRelation=   1
         End
         Begin VB.Label Label2 
            Caption         =   "Líneas Seleccionadas:"
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
            TabIndex        =   36
            Top             =   4080
            Width           =   4092
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   7212
         Left            =   -69400
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   12721
         _StockProps     =   79
         Caption         =   "Consulta:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin VB.ComboBox cboCategoria 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2640
            Width           =   1222
         End
         Begin XtremeSuiteControls.CheckBox chkFechas 
            Height          =   372
            Left            =   1680
            TabIndex        =   5
            Top             =   1540
            Width           =   852
            _Version        =   1441793
            _ExtentX        =   1503
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Todas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Alignment       =   1
         End
         Begin XtremeSuiteControls.RadioButton OptX 
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   2652
            _Version        =   1441793
            _ExtentX        =   4678
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Créditos Formalizados"
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
         Begin XtremeSuiteControls.RadioButton OptX 
            Height          =   372
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   2652
            _Version        =   1441793
            _ExtentX        =   4678
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Créditos Cancelados"
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
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Height          =   732
            Left            =   360
            TabIndex        =   8
            Top             =   6480
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   1291
            _StockProps     =   79
            Caption         =   "Buscar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_Telemarketing_Consultas.frx":2446
            TextImageRelation=   1
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   732
            Left            =   1440
            TabIndex        =   9
            Top             =   6480
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   1291
            _StockProps     =   79
            Caption         =   "Exportar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmAF_Telemarketing_Consultas.frx":2E64
            TextImageRelation=   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpInicio 
            Height          =   312
            Left            =   1320
            TabIndex        =   10
            Top             =   1920
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtpCorte 
            Height          =   312
            Left            =   1320
            TabIndex        =   11
            Top             =   2280
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.CheckBox chkSinMora 
            Height          =   372
            Left            =   360
            TabIndex        =   12
            Top             =   5880
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Operaciones sin Morosidad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit_Linea 
            Height          =   336
            Left            =   1320
            TabIndex        =   13
            Top             =   3000
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
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
         Begin XtremeSuiteControls.CheckBox chkEmail 
            Height          =   372
            Left            =   360
            TabIndex        =   14
            Top             =   5160
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Email Válidos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkMovil 
            Height          =   372
            Left            =   360
            TabIndex        =   15
            Top             =   5520
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Teléfono Móvil Valido"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            TextAlignment   =   2
            Appearance      =   2
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit_Destino 
            Height          =   336
            Left            =   1320
            TabIndex        =   16
            Top             =   3360
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit_Institucion 
            Height          =   336
            Left            =   1320
            TabIndex        =   17
            Top             =   3720
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit_Canal 
            Height          =   336
            Left            =   1320
            TabIndex        =   18
            Top             =   4080
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit_Actividad 
            Height          =   336
            Left            =   1320
            TabIndex        =   19
            Top             =   4440
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
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
         Begin XtremeSuiteControls.RadioButton OptX 
            Height          =   372
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   2652
            _Version        =   1441793
            _ExtentX        =   4678
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Créditos por Finalizar"
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit_Preferencias 
            Height          =   336
            Left            =   1320
            TabIndex        =   44
            Top             =   4800
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "G y P:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   6
            Left            =   240
            TabIndex        =   45
            ToolTipText     =   "Gustos y Preferencias"
            Top             =   4800
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   5
            Left            =   240
            TabIndex        =   42
            Top             =   2640
            Width           =   1212
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Producto:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   4
            Left            =   240
            TabIndex        =   28
            Top             =   4440
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Canal:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   3
            Left            =   240
            TabIndex        =   27
            Top             =   4080
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Institución:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   3720
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Destino:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   3360
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Crédito:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   3000
            Width           =   972
         End
         Begin XtremeSuiteControls.Label lblFecha 
            Height          =   132
            Left            =   120
            TabIndex        =   23
            Top             =   1560
            Width           =   2532
            _Version        =   1441793
            _ExtentX        =   4466
            _ExtentY        =   233
            _StockProps     =   79
            Caption         =   "Fecha Referencia:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   9
            Left            =   360
            TabIndex        =   22
            Top             =   1920
            Width           =   852
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Corte"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   8
            Left            =   360
            TabIndex        =   21
            Top             =   2280
            Width           =   852
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   7332
         Left            =   -66400
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   11052
         _Version        =   524288
         _ExtentX        =   19495
         _ExtentY        =   12933
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
         MaxCols         =   22
         SpreadDesigner  =   "frmAF_Telemarketing_Consultas.frx":3669
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid_Clientes 
         Height          =   3972
         Left            =   -65080
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   11052
         _Version        =   524288
         _ExtentX        =   19495
         _ExtentY        =   7006
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
         MaxCols         =   11
         SpreadDesigner  =   "frmAF_Telemarketing_Consultas.frx":4070
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid_Operaciones 
         Height          =   2772
         Left            =   -65080
         TabIndex        =   40
         Top             =   4560
         Visible         =   0   'False
         Width           =   7572
         _Version        =   524288
         _ExtentX        =   13356
         _ExtentY        =   4890
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
         MaxCols         =   16
         SpreadDesigner  =   "frmAF_Telemarketing_Consultas.frx":4825
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid_Contactos 
         Height          =   3972
         Left            =   4920
         TabIndex        =   46
         Top             =   240
         Width           =   11052
         _Version        =   524288
         _ExtentX        =   19495
         _ExtentY        =   7006
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
         MaxCols         =   11
         SpreadDesigner  =   "frmAF_Telemarketing_Consultas.frx":5100
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
End
Attribute VB_Name = "frmAF_Telemarketing_Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mFecUltMovUpdate As Integer

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnClienteBuscar_Click()
Call sbCliente_Buscar
End Sub

Private Sub btnClienteExportar_Click()
Call sbCliente_Exportar
End Sub

Private Sub btnExportar_Click()
Call sbExportar
End Sub


Private Sub btnPersonas_Click()
Call sbBuscar_Personas

End Sub

Private Sub btnPersonas_Exportar_Click()
Call sbExportar_Personas
End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub

Private Sub sbCliente_Exportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 10
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Apellido No.1"
    vHeaders.Headers(4) = "Apellido No.2"
    vHeaders.Headers(5) = "Correo"
    vHeaders.Headers(6) = "Móvil"
    vHeaders.Headers(7) = "Tel. Hab."
    vHeaders.Headers(8) = "Provincia"
    vHeaders.Headers(9) = "Canton"
    vHeaders.Headers(10) = "Institución"
    
    Call sbSIFGridExportar(vGrid_Clientes, vHeaders, "Telemarketing_ClientesComun")


End Sub


Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 22
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Apellido No.1"
    vHeaders.Headers(4) = "Apellido No.2"
    vHeaders.Headers(5) = "Correo"
    vHeaders.Headers(6) = "Móvil"
    vHeaders.Headers(7) = "Tel. Hab."
    vHeaders.Headers(8) = "Provincia"
    vHeaders.Headers(9) = "Canton"
    vHeaders.Headers(10) = "Línea / Crédito"
    vHeaders.Headers(11) = "Destino"
    vHeaders.Headers(12) = "Producto / Servicio"
    vHeaders.Headers(13) = "Canal"
    vHeaders.Headers(14) = "Monto"
    vHeaders.Headers(15) = "Plazo"
    vHeaders.Headers(16) = "Institución"
    vHeaders.Headers(17) = "Departamento"
    vHeaders.Headers(18) = "Ultimo Movimiento"
    vHeaders.Headers(19) = "Formalización"
    vHeaders.Headers(20) = "Fecha Termina"
    vHeaders.Headers(21) = "Ejecutivo"
    vHeaders.Headers(22) = "Categoría"
    
    Call sbSIFGridExportar(vGrid, vHeaders, "Telemarketing_Consulta")


End Sub



Private Sub sbExportar_Personas()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Apellido No.1"
    vHeaders.Headers(4) = "Apellido No.2"
    vHeaders.Headers(5) = "Email No.1"
    vHeaders.Headers(6) = "Email No.2"
    
    vHeaders.Headers(7) = "Móvil"
    vHeaders.Headers(8) = "Tel. Hab."
    vHeaders.Headers(9) = "Provincia"
    vHeaders.Headers(10) = "Canton"
    vHeaders.Headers(11) = "Institución"
    Call sbSIFGridExportar(vGrid_Contactos, vHeaders, "Telemarketing_Contactos")


End Sub


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

' (@Informe varchar(10), @Inicio datetime, @Corte datetime, @Linea varchar(10), @Mora smallint)"
 
Select Case True
  Case OptX.Item(0).Value 'Formalizados
     strSQL = "exec spAFI_Telemarketing_Consulta 'CRD_01'"
  Case OptX.Item(1).Value 'Cancelados
     strSQL = "exec spAFI_Telemarketing_Consulta 'CRD_02'"
  Case OptX.Item(2).Value 'Finalizacion
     strSQL = "exec spAFI_Telemarketing_Consulta 'CRD_03'"
End Select
 
If chkFechas.Value = xtpChecked Then
  strSQL = strSQL & ",Null, Null"
Else
 strSQL = strSQL & ",'" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
        & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
End If

If Trim(FlatEdit_Linea.Text) = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & FlatEdit_Linea.Text & "'"
End If

If Trim(FlatEdit_Destino.Text) = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & FlatEdit_Destino.Text & "'"
End If

If Trim(FlatEdit_Actividad.Text) = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & FlatEdit_Actividad.Text & "'"
End If

If Trim(FlatEdit_Canal.Text) = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & FlatEdit_Canal.Text & "'"
End If


If Not IsNumeric(FlatEdit_Institucion.Text) Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & "," & FlatEdit_Institucion.Text
End If


strSQL = strSQL & "," & chkSinMora.Value & "," & chkEmail.Value & "," & chkMovil.Value & "," & mFecUltMovUpdate

If cboCategoria.Text = "TODOS" Then
    strSQL = strSQL & ",Null"
Else
 strSQL = strSQL & ",'" & cboCategoria.Text & "'"
End If

If Trim(FlatEdit_Preferencias.Text) = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & FlatEdit_Preferencias.Text & "'"
End If


vGrid.MaxRows = 0

gBox_Procesando.Visible = True
lblProcesando.Caption = "Bajando Información..."
DoEvents

Call OpenRecordSet(rs, strSQL)

lblProcesando.Caption = "Cargando..."
DoEvents

Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 1
  vGrid.Text = rs!Cedula
  vGrid.col = 2
  vGrid.Text = rs!Nombre
  vGrid.col = 3
  vGrid.Text = rs!Apellido1
  vGrid.col = 4
  vGrid.Text = rs!Apellido2
  vGrid.col = 5
  vGrid.Text = rs!Email
  vGrid.col = 6
  vGrid.Text = rs!Movil
  vGrid.col = 7
  vGrid.Text = rs!Tel_Hab
  vGrid.col = 8
  vGrid.Text = rs!Provincia
  vGrid.col = 9
  vGrid.Text = rs!Canton
  vGrid.col = 10
  vGrid.Text = rs!Linea
  vGrid.col = 11
  vGrid.Text = rs!Destino
  
  vGrid.col = 12
  vGrid.Text = rs!Actividad
  
  vGrid.col = 13
  vGrid.Text = rs!Canal
  
  
  vGrid.col = 14
  vGrid.Text = Format(rs!Monto, "Standard")
  vGrid.col = 15
  vGrid.Text = CStr(rs!Plazo)
  vGrid.col = 16
  vGrid.Text = rs!Institucion
  vGrid.col = 17
  vGrid.Text = rs!departamento
  
  
  vGrid.col = 18
  vGrid.Text = Format(rs!Ultimo_Mov & "", "dd/mm/yyyy")
  vGrid.col = 19
  vGrid.Text = Format(rs!FechaForp & "", "dd/mm/yyyy")
  vGrid.col = 20
  vGrid.Text = Format(rs!Fecha_Termina & "", "dd/mm/yyyy")
  
  vGrid.col = 21
  vGrid.Text = rs!Ejecutivo
  
  vGrid.col = 22
  vGrid.Text = rs!Categoria
  
  
  rs.MoveNext
Loop
rs.Close

GroupBox1.Caption = "Consulta (Lineas devueltas: " & Format(vGrid.MaxRows, "###,###,##0") & ")"

Me.MousePointer = vbDefault
gBox_Procesando.Visible = False

mFecUltMovUpdate = 0

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBuscar_Personas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Telemarketing_Contactos "
 
If cboPersonaFecha.Text = "Todas" Then
  strSQL = strSQL & " Null, Null, T"
Else
 strSQL = strSQL & "'" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
        & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00','" & Mid(cboPersonaFecha.Text, 1, 1) & "'"
End If


If cboPersonaEstado.Text = "TODOS" Then
  strSQL = strSQL & ",Null"

Else
  strSQL = strSQL & ",'" & cboPersonaEstado.ItemData(cboPersonaEstado.ListIndex) & "'"
    
End If

vGrid_Contactos.MaxRows = 0

gBox_Procesando.Visible = True
lblProcesando.Caption = "Bajando Información..."
DoEvents

Call OpenRecordSet(rs, strSQL)

lblProcesando.Caption = "Cargando..."
DoEvents

Do While Not rs.EOF
  vGrid_Contactos.MaxRows = vGrid_Contactos.MaxRows + 1
  vGrid_Contactos.Row = vGrid_Contactos.MaxRows
  
  vGrid_Contactos.col = 1
  vGrid_Contactos.Text = rs!Cedula
  vGrid_Contactos.col = 2
  vGrid_Contactos.Text = rs!Nombre
  vGrid_Contactos.col = 3
  vGrid_Contactos.Text = rs!Apellido1
  vGrid_Contactos.col = 4
  vGrid_Contactos.Text = rs!Apellido2
  vGrid_Contactos.col = 5
  vGrid_Contactos.Text = rs!Email
  vGrid_Contactos.col = 6
  vGrid_Contactos.Text = rs!Email_02
  
  
  vGrid_Contactos.col = 7
  vGrid_Contactos.Text = rs!Movil
  vGrid_Contactos.col = 8
  vGrid_Contactos.Text = rs!Tel_Hab
  vGrid_Contactos.col = 8
  vGrid_Contactos.Text = rs!Provincia
  vGrid_Contactos.col = 10
  vGrid_Contactos.Text = rs!Canton
  vGrid_Contactos.col = 11
  
  vGrid_Contactos.col = 12
  vGrid_Contactos.Text = rs!Institucion
  
  
  rs.MoveNext
Loop
rs.Close

GroupBox3.Caption = "Consulta (Lineas devueltas: " & Format(vGrid_Contactos.MaxRows, "###,###,##0") & ")"

Me.MousePointer = vbDefault
gBox_Procesando.Visible = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCliente_Buscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
 
'Registrando Filtros
strSQL = "delete SYS_REPORT_PIVOT_01 where usuario = '" & glogon.Usuario & "'"

For i = 1 To lswSel.ListItems.Count
  strSQL = strSQL & Space(10) & "insert SYS_REPORT_PIVOT_01(USUARIO,CODIGO,REGISTRO_FECHA,COD_REPORTE)" _
         & " VALUES('" & glogon.Usuario & "','" & lswSel.ListItems.Item(i).Text _
         & "',getdate(),'MKD_Clc')"
Next i

Call ConectionExecute(strSQL)


vGrid_Operaciones.MaxRows = 0

strSQL = "exec spMKD_ClientesComun '" & glogon.Usuario & "'," & chkIntegral.Value

With vGrid_Clientes
.MaxRows = 0

gBox_Procesando.Visible = True
lblProcesando.Caption = "Bajando Información..."
DoEvents

Call OpenRecordSet(rs, strSQL)

lblProcesando.Caption = "Cargando..."
DoEvents

vPaso = True
Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  
  .col = 2
  .Text = rs!Cedula
  .col = 3
  .Text = rs!Nombre
  .col = 4
  .Text = rs!Apellido1
  .col = 5
  .Text = rs!Apellido2
  .col = 6
  .Text = rs!Email
  .col = 7
  .Text = rs!Movil
  .col = 8
  .Text = rs!Tel_Hab
  .col = 9
  .Text = rs!Provincia
  .col = 10
  .Text = rs!Canton
  .col = 11
  .Text = rs!Institucion
  
  
  rs.MoveNext
Loop
rs.Close

GroupBox2.Caption = "Consulta (Lineas devueltas: " & Format(.MaxRows, "###,###,##0") & ")"

End With

vPaso = False


Me.MousePointer = vbDefault
gBox_Procesando.Visible = False


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub







Private Sub DateTimePicker2_Change()

End Sub

Private Sub FlatEdit_Actividad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_ACTIVIDAD,DESCRIPCION From AFI_ACTIVIDADES_ECO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Actividad.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Actividad.ToolTipText = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub FlatEdit_Canal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select CANAL_TIPO,DESCRIPCION From AFI_CANALES_TIPOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Canal.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Canal.ToolTipText = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub FlatEdit_Destino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_DESTINO,DESCRIPCION From CATALOGO_DESTINOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Destino.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Destino.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub



Private Sub FlatEdit_Institucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_INSTITUCION,DESCRIPCION From INSTITUCIONES"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Institucion.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Institucion.ToolTipText = Trim(gBusquedas.Resultado2)
End If

End Sub



Private Sub FlatEdit_Preferencias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_PREFERENCIA,DESCRIPCION From AFI_PREFERENCIAS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Preferencias.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Preferencias.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

TabControl_Main.SelectedItem = 0

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -3, dtpCorte.Value)

dtpPersonaInicio.Value = dtpInicio.Value
dtpPersonaCorte.Value = dtpCorte.Value

mFecUltMovUpdate = 1


strSQL = "select COD_ESTADO as 'IdX', DESCRIPCION as 'ItmX' from AFI_ESTADOS_PERSONA "
Call sbCbo_Llena_New(cboPersonaEstado, strSQL, True, True)


cboPersonaFecha.Clear
cboPersonaFecha.AddItem "Ingreso"
cboPersonaFecha.AddItem "Nacimiento"
cboPersonaFecha.Text = "Ingreso"


strSQL = "select COD_MORA as ItmX From Cbr_Clasificacion_Mora"
Call sbLlenaCbo(cboCategoria, strSQL, True, False)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub Form_Resize()

On Error Resume Next

If Me.Height < 7380 Then
   Me.Height = 8000
End If

TabControl_Main.Width = Me.Width - (TabControl_Main.Left + 200)
TabControl_Main.Height = Me.Height - (TabControl_Main.top + 550)

vGrid.Width = TabControl_Main.Width - (120 + vGrid.Left)
vGrid.Height = TabControl_Main.Height - (vGrid.top + 150)

GroupBox1.Height = vGrid.Height

btnBuscar.top = vGrid.Height - btnBuscar.Height - 100
btnExportar.top = btnBuscar.top

chkSinMora.top = btnBuscar.top - 450
chkMovil.top = chkSinMora.top - 360
chkEmail.top = chkMovil.top - 360


GroupBox2.Height = GroupBox1.Height
lswSel.Height = GroupBox2.Height - (lswSel.top + btnClienteBuscar.Height + 100)


vGrid_Clientes.Width = TabControl_Main.Width - (vGrid_Clientes.Left + 180)
vGrid_Operaciones.Width = vGrid_Clientes.Width

vGrid_Clientes.Height = TabControl_Main.Height - (vGrid_Clientes.top + 150 + vGrid_Operaciones.Height)

vGrid_Operaciones.top = vGrid_Clientes.top + vGrid_Clientes.Height + 100

btnClienteBuscar.top = lswSel.Height + lswSel.top + 100
btnClienteExportar.top = btnClienteBuscar.top

chkIntegral.top = btnClienteBuscar.top

GroupBox3.Height = GroupBox1.Height
vGrid_Contactos.Height = vGrid.Height
vGrid_Contactos.Width = TabControl_Main.Width - (vGrid_Contactos.Left + 180)



End Sub

Private Sub FlatEdit_Linea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select CODIGO,DESCRIPCION From CATALOGO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = " and LINEA_INTERNA = 1 AND RETENCION = 'N' AND POLIZA = 'N'"
    frmBusquedas.Show vbModal
    FlatEdit_Linea.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Linea.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub




Private Sub lswCrd_DblClick()
Dim pCodigo As String, pDescripcion As String
Dim itmX As ListItem

If lswCrd.ListItems.Count <= 0 Then Exit Sub

pCodigo = lswCrd.SelectedItem
pDescripcion = lswCrd.SelectedItem.SubItems(1)

Set itmX = lswSel.FindItem(pCodigo, , , lvwWhole)

If itmX Is Nothing Then
   Set itmX = lswSel.ListItems.Add(, , pCodigo)
       itmX.SubItems(1) = pDescripcion
End If



End Sub



Private Sub lswSel_DblClick()

If lswSel.ListItems.Count <= 0 Then Exit Sub

lswSel.ListItems.Remove lswSel.SelectedItem.Index

End Sub

Private Sub OptLineas_Click(Index As Integer)
lswCrd.ListItems.Clear
txtCrd_Consulta.Text = ""
txtCrd_Consulta.SetFocus
Call sbConsultaLineas
End Sub

Private Sub OptX_Click(Index As Integer)
Select Case True
  Case OptX.Item(0).Value
    lblFecha.Caption = "Fecha Formalización:"
  Case OptX.Item(1).Value
    lblFecha.Caption = "Fecha Cancelación:"
  Case OptX.Item(2).Value
    lblFecha.Caption = "Fecha Finalización:"
    
End Select
End Sub


Private Sub sbConsultaLineas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "Select Codigo,Descripcion from Catalogo where (codigo like '%" & txtCrd_Consulta.Text _
        & "%' or descripcion like '%" & txtCrd_Consulta.Text & "%')"

If OptLineas.Item(0).Value = True Then
    strSQL = strSQL & " and Poliza = 'N' and Retencion = 'N'"
Else
    strSQL = strSQL & " and Poliza = 'S' or Retencion = 'S'"
End If

lswCrd.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswCrd.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!DESCRIPCION
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub TabControl_Main_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 Call Form_Resize
End Sub

Private Sub txtCrd_Consulta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbConsultaLineas

End Sub

Private Sub vGrid_Clientes_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

vGrid_Operaciones.MaxRows = 0


Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
 
vGrid_Clientes.Row = Row
vGrid_Clientes.col = 2
vCadena = vGrid_Clientes.Text

With vGrid_Operaciones

.MaxRows = 0

strSQL = "exec spMKD_ClientesComun_Detalle '" & vCadena & "','" & glogon.Usuario & "'"


gBox_Procesando.Visible = True
lblProcesando.Caption = "Bajando Información..."
DoEvents

Call OpenRecordSet(rs, strSQL)

lblProcesando.Caption = "Cargando..."
DoEvents

vPaso = True
Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  .col = 1
  .Text = CStr(rs!Id_Solicitud)
  .col = 2
  .Text = rs!Codigo
  .col = 3
  .Text = rs!Linea
  .col = 4
  .Text = rs!Destino
  .col = 5
  .Text = rs!Actividad
  .col = 6
  .Text = rs!Canal

  .col = 7
  .Text = Format(rs!Monto, "Standard")
  .col = 8
  .Text = CStr(rs!Plazo)
  .col = 9
  .Text = Format(rs!Tasa, "Standard")
  .col = 10
  .Text = Format(rs!Cuota, "Standard")
  .col = 11
  .Text = Format(rs!Saldo, "Standard")
  
  .col = 12
  .Text = rs!Institucion
  
  .col = 13
  .Text = Format(rs!Ultimo_Mov & "", "dd/mm/yyyy")
  .col = 14
  .Text = Format(rs!FechaForp & "", "dd/mm/yyyy")
  .col = 15
  .Text = Format(rs!Fecha_Termina & "", "dd/mm/yyyy")
  
  .col = 16
  .Text = rs!Ejecutivo
  
  rs.MoveNext
Loop
rs.Close

GroupBox2.Caption = "Consulta (Lineas devueltas: " & Format(.MaxRows, "###,###,##0") & ")"

End With

vPaso = False


Me.MousePointer = vbDefault
gBox_Procesando.Visible = False


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub
