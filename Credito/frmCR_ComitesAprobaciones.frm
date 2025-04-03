VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_ComitesAprobaciones 
   Caption         =   "Aprobación de Comites"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   16080
   WindowState     =   2  'Maximized
   Begin VB.Frame fraActa 
      BorderStyle     =   0  'None
      Caption         =   "Mantenimiento del Acta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   10440
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   11412
      Begin XtremeSuiteControls.TabControl tcActas 
         Height          =   6615
         Left            =   0
         TabIndex        =   151
         Top             =   120
         Width           =   11415
         _Version        =   1572864
         _ExtentX        =   20135
         _ExtentY        =   11668
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
         Item(0).Caption =   "Actual"
         Item(0).ControlCount=   15
         Item(0).Control(0)=   "lswAsistencia"
         Item(0).Control(1)=   "btnActaTool(0)"
         Item(0).Control(2)=   "btnActaTool(1)"
         Item(0).Control(3)=   "dtpActaFecha"
         Item(0).Control(4)=   "txtActasNotas"
         Item(0).Control(5)=   "LabelX(2)"
         Item(0).Control(6)=   "LabelX(1)"
         Item(0).Control(7)=   "LabelX(0)"
         Item(0).Control(8)=   "lblUsuario(5)"
         Item(0).Control(9)=   "LabelX(4)"
         Item(0).Control(10)=   "txtSesion"
         Item(0).Control(11)=   "LabelX(3)"
         Item(0).Control(12)=   "txtActa"
         Item(0).Control(13)=   "txtActaEstado"
         Item(0).Control(14)=   "btnActaTool(2)"
         Item(1).Caption =   "Histórico"
         Item(1).ControlCount=   8
         Item(1).Control(0)=   "dtpActaInicio"
         Item(1).Control(1)=   "dtpActaCorte"
         Item(1).Control(2)=   "Label2"
         Item(1).Control(3)=   "lswActaH"
         Item(1).Control(4)=   "lblUsuario(4)"
         Item(1).Control(5)=   "btnActaConsulta"
         Item(1).Control(6)=   "txtActaFiltro"
         Item(1).Control(7)=   "btnExportar(0)"
         Item(2).Caption =   "Resoluciones"
         Item(2).ControlCount=   3
         Item(2).Control(0)=   "lswActaR"
         Item(2).Control(1)=   "btnExportar(1)"
         Item(2).Control(2)=   "ShortcutCaption1"
         Begin XtremeSuiteControls.ListView lswActaR 
            Height          =   5655
            Left            =   -70000
            TabIndex        =   176
            Top             =   840
            Visible         =   0   'False
            Width           =   11415
            _Version        =   1572864
            _ExtentX        =   20135
            _ExtentY        =   9975
            _StockProps     =   77
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
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.ListView lswAsistencia 
            Height          =   5655
            Left            =   4800
            TabIndex        =   152
            Top             =   840
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
            _ExtentY        =   9975
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
            Checkboxes      =   -1  'True
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswActaH 
            Height          =   5535
            Left            =   -70000
            TabIndex        =   167
            Top             =   1080
            Visible         =   0   'False
            Width           =   11415
            _Version        =   1572864
            _ExtentX        =   20135
            _ExtentY        =   9763
            _StockProps     =   77
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
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnActaTool 
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   153
            Top             =   840
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Nueva"
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
         Begin XtremeSuiteControls.PushButton btnActaTool 
            Height          =   555
            Index           =   1
            Left            =   3240
            TabIndex        =   154
            Top             =   5520
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   979
            _StockProps     =   79
            Caption         =   "Guardar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":0000
         End
         Begin XtremeSuiteControls.DateTimePicker dtpActaFecha 
            Height          =   315
            Left            =   1080
            TabIndex        =   155
            Top             =   1800
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2773
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
         Begin XtremeSuiteControls.FlatEdit txtActasNotas 
            Height          =   2715
            Left            =   1080
            TabIndex        =   156
            Top             =   2640
            Width           =   3615
            _Version        =   1572864
            _ExtentX        =   6371
            _ExtentY        =   4784
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
         Begin XtremeSuiteControls.FlatEdit txtSesion 
            Height          =   315
            Left            =   1080
            TabIndex        =   162
            Top             =   1200
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   550
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpActaInicio 
            Height          =   315
            Left            =   -69040
            TabIndex        =   164
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   550
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtpActaCorte 
            Height          =   315
            Left            =   -67720
            TabIndex        =   165
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   550
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.FlatEdit txtActaFiltro 
            Height          =   315
            Left            =   -64240
            TabIndex        =   168
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnActaConsulta 
            Height          =   435
            Left            =   -62080
            TabIndex        =   171
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Buscar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":0731
         End
         Begin XtremeSuiteControls.FlatEdit txtActa 
            Height          =   315
            Left            =   1080
            TabIndex        =   172
            Top             =   840
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   550
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtActaEstado 
            Height          =   315
            Left            =   1080
            TabIndex        =   174
            Top             =   2280
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   550
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnActaTool 
            Height          =   315
            Index           =   2
            Left            =   2640
            TabIndex        =   175
            Top             =   2280
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Cerrar"
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
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   435
            Index           =   0
            Left            =   -60760
            TabIndex        =   177
            Top             =   480
            Visible         =   0   'False
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   767
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":0E31
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   435
            Index           =   1
            Left            =   -59080
            TabIndex        =   178
            Top             =   360
            Visible         =   0   'False
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   767
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":0F9B
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   -70000
            TabIndex        =   179
            Top             =   360
            Visible         =   0   'False
            Width           =   11295
            _Version        =   1572864
            _ExtentX        =   19923
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Lista de Casos con Resolución dentro del Acta:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblUsuario 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Identificación"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   -65920
            TabIndex        =   169
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fechas:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -69760
            TabIndex        =   166
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LabelX 
            BackStyle       =   0  'Transparent
            Caption         =   "Sesión:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   163
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label LabelX 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Acta:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   161
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblUsuario 
            BackStyle       =   0  'Transparent
            Caption         =   "Asistencia:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   4800
            TabIndex        =   160
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label LabelX 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   159
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label LabelX 
            BackStyle       =   0  'Transparent
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   158
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label LabelX 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   157
            Top             =   2640
            Width           =   1095
         End
      End
   End
   Begin VB.Frame fraCausas 
      Caption         =   "Causas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   13920
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   12612
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5052
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   11052
         _Version        =   1572864
         _ExtentX        =   19494
         _ExtentY        =   8911
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnCausas 
         Height          =   312
         Left            =   9960
         TabIndex        =   131
         Top             =   480
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   6
         Picture         =   "frmCR_ComitesAprobaciones.frx":1105
         ImageAlignment  =   0
      End
      Begin VB.Label Label22 
         Caption         =   "Indique las causas por las cuales esta indicando que la solicitud queda Pendiente o Denegada con las opciones siguientes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   7935
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   240
         Picture         =   "frmCR_ComitesAprobaciones.frx":181B
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame FraControles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   15255
      Begin XtremeSuiteControls.TabControl tcMain 
         Height          =   3615
         Left            =   0
         TabIndex        =   37
         Top             =   600
         Width           =   14655
         _Version        =   1572864
         _ExtentX        =   25850
         _ExtentY        =   6376
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
         ItemCount       =   5
         SelectedItem    =   3
         Item(0).Caption =   "Detalle"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "tcDetalle"
         Item(1).Caption =   "Fiadores"
         Item(1).ControlCount=   23
         Item(1).Control(0)=   "lswFiadores"
         Item(1).Control(1)=   "Label32"
         Item(1).Control(2)=   "Label28"
         Item(1).Control(3)=   "lblFiadorLiqSFianzaPorc"
         Item(1).Control(4)=   "lblFiadorLiqCFianzaPorc"
         Item(1).Control(5)=   "lblFiadorInstitucion"
         Item(1).Control(6)=   "Label27"
         Item(1).Control(7)=   "Label78"
         Item(1).Control(8)=   "Label77"
         Item(1).Control(9)=   "Label76"
         Item(1).Control(10)=   "lblFiadorLiqSFianza"
         Item(1).Control(11)=   "lblFiadorLiqCFianza"
         Item(1).Control(12)=   "lblFiadorSalLiquido"
         Item(1).Control(13)=   "Label72"
         Item(1).Control(14)=   "lblFLugarTrabajo"
         Item(1).Control(15)=   "lblFiadorIngreso"
         Item(1).Control(16)=   "Label67"
         Item(1).Control(17)=   "lblFiadorNombramiento"
         Item(1).Control(18)=   "Label65"
         Item(1).Control(19)=   "lblFiadorEstado"
         Item(1).Control(20)=   "lblFiadorMembresia"
         Item(1).Control(21)=   "Label62"
         Item(1).Control(22)=   "Label61"
         Item(2).Caption =   "Seguimiento"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "vGridSeguimiento"
         Item(3).Caption =   "Resolución"
         Item(3).ControlCount=   9
         Item(3).Control(0)=   "txtObservacion"
         Item(3).Control(1)=   "Label8(0)"
         Item(3).Control(2)=   "btnResolucion(0)"
         Item(3).Control(3)=   "btnResolucion(1)"
         Item(3).Control(4)=   "btnResolucion(2)"
         Item(3).Control(5)=   "Label8(1)"
         Item(3).Control(6)=   "txtAcuerdoJD"
         Item(3).Control(7)=   "btnResolucion(3)"
         Item(3).Control(8)=   "btnResolucion(4)"
         Item(4).Caption =   "Causas"
         Item(4).ControlCount=   3
         Item(4).Control(0)=   "lswCausasList"
         Item(4).Control(1)=   "optCausas(0)"
         Item(4).Control(2)=   "optCausas(1)"
         Begin XtremeSuiteControls.ListView lswCausasList 
            Height          =   3132
            Left            =   -67240
            TabIndex        =   148
            Top             =   480
            Visible         =   0   'False
            Width           =   9252
            _Version        =   1572864
            _ExtentX        =   16319
            _ExtentY        =   5524
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
         Begin XtremeSuiteControls.ListView lswFiadores 
            Height          =   2892
            Left            =   -69880
            TabIndex        =   38
            Top             =   480
            Visible         =   0   'False
            Width           =   6372
            _Version        =   1572864
            _ExtentX        =   11239
            _ExtentY        =   5101
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
            ShowBorder      =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnResolucion 
            Height          =   615
            Index           =   0
            Left            =   1440
            TabIndex        =   123
            Top             =   2520
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Aprobar   "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":21A3
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.TabControl tcDetalle 
            Height          =   3615
            Left            =   -70000
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   14295
            _Version        =   1572864
            _ExtentX        =   25215
            _ExtentY        =   6376
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
            ItemCount       =   7
            Item(0).Caption =   "Crédito"
            Item(0).ControlCount=   31
            Item(0).Control(0)=   "Label5"
            Item(0).Control(1)=   "lblMembresia"
            Item(0).Control(2)=   "Label7"
            Item(0).Control(3)=   "lblEstadoLaboral"
            Item(0).Control(4)=   "Label11"
            Item(0).Control(5)=   "lblEstadoActual"
            Item(0).Control(6)=   "Label30"
            Item(0).Control(7)=   "lblLugarTrabajo"
            Item(0).Control(8)=   "Label15"
            Item(0).Control(9)=   "Label14"
            Item(0).Control(10)=   "Label13"
            Item(0).Control(11)=   "lblTotalCuotas"
            Item(0).Control(12)=   "lblCuotaDesembolsos"
            Item(0).Control(13)=   "lblCuotaRefundicion"
            Item(0).Control(14)=   "lbl"
            Item(0).Control(15)=   "lblMontoDesembolsos"
            Item(0).Control(16)=   "lblMontoRefundicion"
            Item(0).Control(17)=   "lblMonto_Girado"
            Item(0).Control(18)=   "lblMontoApr"
            Item(0).Control(19)=   "Label10"
            Item(0).Control(20)=   "Label9"
            Item(0).Control(21)=   "Label6"
            Item(0).Control(22)=   "Label26"
            Item(0).Control(23)=   "lblDiferenciaCuota"
            Item(0).Control(24)=   "lblCuota"
            Item(0).Control(25)=   "Label17"
            Item(0).Control(26)=   "Label16(0)"
            Item(0).Control(27)=   "Label16(1)"
            Item(0).Control(28)=   "lblLinea"
            Item(0).Control(29)=   "lblCA"
            Item(0).Control(30)=   "Label4"
            Item(1).Caption =   "Clasificación"
            Item(1).ControlCount=   2
            Item(1).Control(0)=   "vGrid"
            Item(1).Control(1)=   "lblClasificacion"
            Item(2).Caption =   "Patrimonio"
            Item(2).ControlCount=   22
            Item(2).Control(0)=   "Label21"
            Item(2).Control(1)=   "Label20(0)"
            Item(2).Control(2)=   "Label19"
            Item(2).Control(3)=   "Label18"
            Item(2).Control(4)=   "Label12"
            Item(2).Control(5)=   "Label44(0)"
            Item(2).Control(6)=   "Label42"
            Item(2).Control(7)=   "Label39"
            Item(2).Control(8)=   "txtPatrimonio"
            Item(2).Control(9)=   "txtPAT_Disponible"
            Item(2).Control(10)=   "txtPAT_Saldos"
            Item(2).Control(11)=   "txtAhorro"
            Item(2).Control(12)=   "txtAporte"
            Item(2).Control(13)=   "txtCapitalizacion"
            Item(2).Control(14)=   "txtCustodia"
            Item(2).Control(15)=   "lblCapitalizado"
            Item(2).Control(16)=   "lblFechaCustodia"
            Item(2).Control(17)=   "lblFechaAporte"
            Item(2).Control(18)=   "lblFechaAhorro"
            Item(2).Control(19)=   "txtPAT_DisponibleBruto"
            Item(2).Control(20)=   "Label20(1)"
            Item(2).Control(21)=   "txtFondos"
            Item(3).Caption =   "Deudas"
            Item(3).ControlCount=   8
            Item(3).Control(0)=   "vGridDeudas"
            Item(3).Control(1)=   "Label25"
            Item(3).Control(2)=   "lblDeudasCuota"
            Item(3).Control(3)=   "lblDeducciones"
            Item(3).Control(4)=   "lblDeudasTotal"
            Item(3).Control(5)=   "Label34"
            Item(3).Control(6)=   "Label33"
            Item(3).Control(7)=   "scTitulos(3)"
            Item(4).Caption =   "Fianzas"
            Item(4).ControlCount=   8
            Item(4).Control(0)=   "vGridFianzas"
            Item(4).Control(1)=   "lblFianzasMonto"
            Item(4).Control(2)=   "Label29"
            Item(4).Control(3)=   "lblFianzasSaldo"
            Item(4).Control(4)=   "Label31"
            Item(4).Control(5)=   "lblFianzasCuota"
            Item(4).Control(6)=   "Label51"
            Item(4).Control(7)=   "scTitulos(2)"
            Item(5).Caption =   "Refundiciones"
            Item(5).ControlCount=   6
            Item(5).Control(0)=   "vGridRefundiciones"
            Item(5).Control(1)=   "lblRefundeCuota"
            Item(5).Control(2)=   "Label43"
            Item(5).Control(3)=   "Label45"
            Item(5).Control(4)=   "lblRefundeMonto"
            Item(5).Control(5)=   "scTitulos(1)"
            Item(6).Caption =   "Desembolsos"
            Item(6).ControlCount=   6
            Item(6).Control(0)=   "vGridDesembolsos"
            Item(6).Control(1)=   "Label49"
            Item(6).Control(2)=   "lblDesembolsoCuota"
            Item(6).Control(3)=   "lblDesembolsoMonto"
            Item(6).Control(4)=   "Label47"
            Item(6).Control(5)=   "scTitulos(0)"
            Begin FPSpreadADO.fpSpread vGrid 
               Height          =   2055
               Left            =   -65800
               TabIndex        =   40
               Top             =   600
               Visible         =   0   'False
               Width           =   7215
               _Version        =   524288
               _ExtentX        =   12721
               _ExtentY        =   3620
               _StockProps     =   64
               BorderStyle     =   0
               DisplayRowHeaders=   0   'False
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ScrollBars      =   0
               SpreadDesigner  =   "frmCR_ComitesAprobaciones.frx":2915
               AppearanceStyle =   1
            End
            Begin FPSpreadADO.fpSpread vGridFianzas 
               Height          =   2652
               Left            =   -67360
               TabIndex        =   41
               Top             =   480
               Visible         =   0   'False
               Width           =   10692
               _Version        =   524288
               _ExtentX        =   18860
               _ExtentY        =   4678
               _StockProps     =   64
               BorderStyle     =   0
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   499
               SpreadDesigner  =   "frmCR_ComitesAprobaciones.frx":2EBD
               AppearanceStyle =   1
            End
            Begin FPSpreadADO.fpSpread vGridDesembolsos 
               Height          =   2652
               Left            =   -67360
               TabIndex        =   42
               Top             =   480
               Visible         =   0   'False
               Width           =   10452
               _Version        =   524288
               _ExtentX        =   18436
               _ExtentY        =   4678
               _StockProps     =   64
               BorderStyle     =   0
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   492
               SpreadDesigner  =   "frmCR_ComitesAprobaciones.frx":3699
               AppearanceStyle =   1
            End
            Begin FPSpreadADO.fpSpread vGridDeudas 
               Height          =   2892
               Left            =   -67360
               TabIndex        =   132
               Top             =   480
               Visible         =   0   'False
               Width           =   9132
               _Version        =   524288
               _ExtentX        =   16108
               _ExtentY        =   5101
               _StockProps     =   64
               BackColorStyle  =   1
               BorderStyle     =   0
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   17
               SpreadDesigner  =   "frmCR_ComitesAprobaciones.frx":3C70
               VScrollSpecialType=   2
               AppearanceStyle =   1
            End
            Begin XtremeSuiteControls.FlatEdit txtPatrimonio 
               Height          =   312
               Left            =   -60400
               TabIndex        =   133
               Top             =   720
               Visible         =   0   'False
               Width           =   1572
               _Version        =   1572864
               _ExtentX        =   2773
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPAT_Disponible 
               Height          =   312
               Left            =   -60400
               TabIndex        =   134
               Top             =   2280
               Visible         =   0   'False
               Width           =   1572
               _Version        =   1572864
               _ExtentX        =   2773
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPAT_Saldos 
               Height          =   312
               Left            =   -60400
               TabIndex        =   135
               Top             =   1680
               Visible         =   0   'False
               Width           =   1572
               _Version        =   1572864
               _ExtentX        =   2773
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAhorro 
               Height          =   312
               Left            =   -66160
               TabIndex        =   136
               Top             =   720
               Visible         =   0   'False
               Width           =   1452
               _Version        =   1572864
               _ExtentX        =   2561
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAporte 
               Height          =   312
               Left            =   -66160
               TabIndex        =   137
               Top             =   1080
               Visible         =   0   'False
               Width           =   1452
               _Version        =   1572864
               _ExtentX        =   2561
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCapitalizacion 
               Height          =   312
               Left            =   -66160
               TabIndex        =   138
               Top             =   1440
               Visible         =   0   'False
               Width           =   1452
               _Version        =   1572864
               _ExtentX        =   2561
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCustodia 
               Height          =   312
               Left            =   -66160
               TabIndex        =   139
               Top             =   1800
               Visible         =   0   'False
               Width           =   1452
               _Version        =   1572864
               _ExtentX        =   2561
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPAT_DisponibleBruto 
               Height          =   312
               Left            =   -60400
               TabIndex        =   144
               Top             =   1320
               Visible         =   0   'False
               Width           =   1572
               _Version        =   1572864
               _ExtentX        =   2773
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtFondos 
               Height          =   312
               Left            =   -66160
               TabIndex        =   146
               Top             =   2160
               Visible         =   0   'False
               Width           =   1452
               _Version        =   1572864
               _ExtentX        =   2561
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin FPSpreadADO.fpSpread vGridRefundiciones 
               Height          =   2772
               Left            =   -67360
               TabIndex        =   147
               Top             =   480
               Visible         =   0   'False
               Width           =   9132
               _Version        =   524288
               _ExtentX        =   16108
               _ExtentY        =   4890
               _StockProps     =   64
               BackColorStyle  =   1
               BorderStyle     =   0
               DisplayRowHeaders=   0   'False
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
               MaxCols         =   13
               SpreadDesigner  =   "frmCR_ComitesAprobaciones.frx":4ABC
               VScrollSpecialType=   2
               AppearanceStyle =   1
            End
            Begin XtremeSuiteControls.Label lblClasificacion 
               Height          =   375
               Left            =   -69640
               TabIndex        =   190
               Top             =   720
               Visible         =   0   'False
               Width           =   3135
               _Version        =   1572864
               _ExtentX        =   5530
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Clasificación de la Persona:         A"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "( + ) Salario Variable"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   11160
               TabIndex        =   189
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label lblCA 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   11160
               TabIndex        =   188
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label lblLinea 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   9240
               TabIndex        =   186
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Línea de Crédito"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   9240
               TabIndex        =   185
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Ahorros Extraordinarios"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   1
               Left            =   -68320
               TabIndex        =   145
               Top             =   2160
               Visible         =   0   'False
               Width           =   2052
            End
            Begin VB.Label lblFechaAhorro 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "10-1998"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   312
               Left            =   -64720
               TabIndex        =   143
               ToolTipText     =   "Fecha del último ahorro obrero reportado"
               Top             =   720
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.Label lblFechaAporte 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "10-1998"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   312
               Left            =   -64720
               TabIndex        =   142
               ToolTipText     =   "Fecha del último aporte patronal reportado"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.Label lblFechaCustodia 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "10-1998"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   312
               Left            =   -64720
               TabIndex        =   141
               ToolTipText     =   "Fecha del último ahorro extraordinario de este socio"
               Top             =   1800
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.Label lblCapitalizado 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "09-1997"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   312
               Left            =   -64720
               TabIndex        =   140
               ToolTipText     =   "Fecha de la capitalización de los excedentes"
               Top             =   1440
               Visible         =   0   'False
               Width           =   1212
            End
            Begin XtremeShortcutBar.ShortcutCaption scTitulos 
               Height          =   312
               Index           =   3
               Left            =   -70000
               TabIndex        =   129
               Top             =   480
               Visible         =   0   'False
               Width           =   2652
               _Version        =   1572864
               _ExtentX        =   4678
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Deudas y Otros Rebajos"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               Alignment       =   1
            End
            Begin XtremeShortcutBar.ShortcutCaption scTitulos 
               Height          =   312
               Index           =   2
               Left            =   -70000
               TabIndex        =   128
               Top             =   480
               Visible         =   0   'False
               Width           =   2652
               _Version        =   1572864
               _ExtentX        =   4678
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Fianzas"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               Alignment       =   1
            End
            Begin XtremeShortcutBar.ShortcutCaption scTitulos 
               Height          =   312
               Index           =   1
               Left            =   -70000
               TabIndex        =   127
               Top             =   480
               Visible         =   0   'False
               Width           =   2652
               _Version        =   1572864
               _ExtentX        =   4678
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Refundiciones"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               Alignment       =   1
            End
            Begin XtremeShortcutBar.ShortcutCaption scTitulos 
               Height          =   315
               Index           =   0
               Left            =   -70000
               TabIndex        =   126
               Top             =   480
               Visible         =   0   'False
               Width           =   2652
               _Version        =   1572864
               _ExtentX        =   4678
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Desembolsos"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               Alignment       =   1
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Membresía"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   240
               TabIndex        =   97
               Top             =   480
               Width           =   1092
            End
            Begin VB.Label lblMembresia 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   372
               Left            =   240
               TabIndex        =   96
               Top             =   780
               Width           =   3132
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Nombramiento"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   240
               TabIndex        =   95
               Top             =   1200
               Width           =   1332
            End
            Begin VB.Label lblEstadoLaboral 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   372
               Left            =   240
               TabIndex        =   94
               Top             =   1500
               Width           =   1572
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Estado"
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
               Left            =   1920
               TabIndex        =   93
               Top             =   1200
               Width           =   1092
            End
            Begin VB.Label lblEstadoActual 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   372
               Left            =   1920
               TabIndex        =   92
               Top             =   1500
               Width           =   1452
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "Lugar Trabajo"
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
               TabIndex        =   91
               Top             =   1920
               Width           =   1452
            End
            Begin VB.Label lblLugarTrabajo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   240
               TabIndex        =   90
               Top             =   2220
               Width           =   3135
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Desembolsos"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   3960
               TabIndex        =   89
               Top             =   1920
               Width           =   1452
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Refundición"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   3960
               TabIndex        =   88
               Top             =   1560
               Width           =   1452
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Monto Girado"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   3960
               TabIndex        =   87
               Top             =   1200
               Width           =   1452
            End
            Begin VB.Label lblTotalCuotas 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   7200
               TabIndex        =   86
               Top             =   2400
               Width           =   1800
            End
            Begin VB.Label lblCuotaDesembolsos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   7200
               TabIndex        =   85
               Top             =   1920
               Width           =   1800
            End
            Begin VB.Label lblCuotaRefundicion 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   7200
               TabIndex        =   84
               Top             =   1560
               Width           =   1800
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   7200
               TabIndex        =   83
               Top             =   840
               Width           =   1800
            End
            Begin VB.Label lblMontoDesembolsos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   5280
               TabIndex        =   82
               Top             =   1920
               Width           =   1800
            End
            Begin VB.Label lblMontoRefundicion 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   5280
               TabIndex        =   81
               Top             =   1560
               Width           =   1800
            End
            Begin VB.Label lblMonto_Girado 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   5280
               TabIndex        =   80
               Top             =   1200
               Width           =   1800
            End
            Begin VB.Label lblMontoApr 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   5280
               TabIndex        =   79
               Top             =   840
               Width           =   1800
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Monto Solicitado"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   3960
               TabIndex        =   78
               Top             =   840
               Width           =   1212
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Cuota"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   7440
               TabIndex        =   77
               Top             =   480
               Width           =   1092
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Monto"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   5520
               TabIndex        =   76
               Top             =   480
               Width           =   1092
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               Caption         =   "Total Cuotas Liberadas: "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   4560
               TabIndex        =   75
               Top             =   2400
               Width           =   2532
            End
            Begin VB.Label lblDiferenciaCuota 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   9240
               TabIndex        =   74
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label lblCuota 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   9240
               TabIndex        =   73
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Aumenta/Disminuye"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   9240
               TabIndex        =   72
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Nueva Cuota"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   9240
               TabIndex        =   71
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Patrimonio Total"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -62440
               TabIndex        =   70
               Top             =   720
               Visible         =   0   'False
               Width           =   1932
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Patronal en Custodia"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   0
               Left            =   -68320
               TabIndex        =   69
               Top             =   1800
               Visible         =   0   'False
               Width           =   1812
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Patronal"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -68320
               TabIndex        =   68
               Top             =   1080
               Visible         =   0   'False
               Width           =   1692
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Obrero"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -68320
               TabIndex        =   67
               Top             =   720
               Visible         =   0   'False
               Width           =   1572
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Capitalización"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -68320
               TabIndex        =   66
               Top             =   1440
               Visible         =   0   'False
               Width           =   1452
            End
            Begin VB.Label Label44 
               BackStyle       =   0  'Transparent
               Caption         =   "Disponible"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   0
               Left            =   -62440
               TabIndex        =   65
               Top             =   2280
               Visible         =   0   'False
               Width           =   1932
            End
            Begin VB.Label Label42 
               BackStyle       =   0  'Transparent
               Caption         =   "(-) Saldo préstamos"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -62440
               TabIndex        =   64
               Top             =   1680
               Visible         =   0   'False
               Width           =   1452
            End
            Begin VB.Label Label39 
               BackStyle       =   0  'Transparent
               Caption         =   "Disponible bruto"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -62440
               TabIndex        =   63
               Top             =   1320
               Visible         =   0   'False
               Width           =   1452
            End
            Begin VB.Label Label25 
               Caption         =   "Total Cuota"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69640
               TabIndex        =   62
               Top             =   1560
               Visible         =   0   'False
               Width           =   1932
            End
            Begin VB.Label lblDeudasCuota 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69640
               TabIndex        =   61
               Top             =   1800
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label lblDeducciones 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69640
               TabIndex        =   60
               Top             =   2400
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label lblDeudasTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69640
               TabIndex        =   59
               Top             =   1200
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label Label34 
               Caption         =   "Deducciones"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69640
               TabIndex        =   58
               Top             =   2160
               Visible         =   0   'False
               Width           =   1332
            End
            Begin VB.Label Label33 
               Caption         =   "Total Saldo"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69640
               TabIndex        =   57
               Top             =   960
               Visible         =   0   'False
               Width           =   852
            End
            Begin VB.Label lblFianzasMonto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69520
               TabIndex        =   56
               Top             =   1200
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label Label29 
               Caption         =   "Monto"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69520
               TabIndex        =   55
               Top             =   960
               Visible         =   0   'False
               Width           =   852
            End
            Begin VB.Label lblFianzasSaldo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69520
               TabIndex        =   54
               Top             =   1800
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label Label31 
               Caption         =   "Saldo"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69520
               TabIndex        =   53
               Top             =   1560
               Visible         =   0   'False
               Width           =   852
            End
            Begin VB.Label lblFianzasCuota 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69520
               TabIndex        =   52
               Top             =   2400
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label Label51 
               Caption         =   "Cuota"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69520
               TabIndex        =   51
               Top             =   2160
               Visible         =   0   'False
               Width           =   852
            End
            Begin VB.Label lblRefundeCuota 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69520
               TabIndex        =   50
               Top             =   1920
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label Label43 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuota + Póliza Liberada"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69520
               TabIndex        =   49
               Top             =   1680
               Visible         =   0   'False
               Width           =   2052
            End
            Begin VB.Label Label45 
               BackStyle       =   0  'Transparent
               Caption         =   "Monto"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69520
               TabIndex        =   48
               Top             =   1080
               Visible         =   0   'False
               Width           =   852
            End
            Begin VB.Label lblRefundeMonto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69520
               TabIndex        =   47
               Top             =   1320
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label Label49 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuota Liberada"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69520
               TabIndex        =   46
               Top             =   1560
               Visible         =   0   'False
               Width           =   2052
            End
            Begin VB.Label lblDesembolsoCuota 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69520
               TabIndex        =   45
               Top             =   1800
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label lblDesembolsoMonto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   -69520
               TabIndex        =   44
               Top             =   1200
               Visible         =   0   'False
               Width           =   1992
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "Monto"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   -69520
               TabIndex        =   43
               Top             =   960
               Visible         =   0   'False
               Width           =   1092
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtObservacion 
            Height          =   1635
            Left            =   1440
            TabIndex        =   98
            Top             =   480
            Width           =   9615
            _Version        =   1572864
            _ExtentX        =   16960
            _ExtentY        =   2884
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
         Begin FPSpreadADO.fpSpread vGridSeguimiento 
            Height          =   3252
            Left            =   -69880
            TabIndex        =   99
            Top             =   360
            Visible         =   0   'False
            Width           =   10932
            _Version        =   524288
            _ExtentX        =   19283
            _ExtentY        =   5736
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
            SpreadDesigner  =   "frmCR_ComitesAprobaciones.frx":54B1
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton btnResolucion 
            Height          =   615
            Index           =   1
            Left            =   3240
            TabIndex        =   124
            Top             =   2520
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Pendiente"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":5A47
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnResolucion 
            Height          =   615
            Index           =   2
            Left            =   9120
            TabIndex        =   125
            Top             =   2520
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Denegar   "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":61B9
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.RadioButton optCausas 
            Height          =   492
            Index           =   0
            Left            =   -69760
            TabIndex        =   149
            Top             =   480
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1572864
            _ExtentX        =   4043
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Causas para Denegación"
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optCausas 
            Height          =   492
            Index           =   1
            Left            =   -69760
            TabIndex        =   150
            Top             =   960
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1572864
            _ExtentX        =   4043
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Pendientes para Estudio"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtAcuerdoJD 
            Height          =   315
            Left            =   1440
            TabIndex        =   182
            Top             =   2160
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   556
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnResolucion 
            Height          =   615
            Index           =   3
            Left            =   5040
            TabIndex        =   183
            Top             =   2520
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "VB %Liq Pen.Avalúo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":692B
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnResolucion 
            Height          =   615
            Index           =   4
            Left            =   7080
            TabIndex        =   184
            Top             =   2520
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "VB %Liq. Pend. Acuerdo/Factura"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmCR_ComitesAprobaciones.frx":D18D
            ImageAlignment  =   4
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Acuerdo JD"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   181
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Observación"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   122
            Top             =   480
            Width           =   1452
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -59200
            TabIndex        =   121
            Top             =   2760
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -61120
            TabIndex        =   120
            Top             =   2760
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Label lblFiadorLiqSFianzaPorc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -58360
            TabIndex        =   119
            ToolTipText     =   "Membresia"
            Top             =   3000
            Visible         =   0   'False
            Width           =   888
         End
         Begin VB.Label lblFiadorLiqCFianzaPorc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -61480
            TabIndex        =   118
            ToolTipText     =   "Membresia"
            Top             =   3000
            Visible         =   0   'False
            Width           =   888
         End
         Begin VB.Label lblFiadorInstitucion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -63400
            TabIndex        =   117
            ToolTipText     =   "Membresia"
            Top             =   1800
            Visible         =   0   'False
            Width           =   5892
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Institución"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -63400
            TabIndex        =   116
            Top             =   1560
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label Label78 
            BackStyle       =   0  'Transparent
            Caption         =   "Liq S/Fianzas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -60400
            TabIndex        =   115
            Top             =   2760
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label Label77 
            BackStyle       =   0  'Transparent
            Caption         =   "Liq C/Fianzas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -63400
            TabIndex        =   114
            Top             =   2760
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "Salario Liquido"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -59440
            TabIndex        =   113
            Top             =   2160
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label lblFiadorLiqSFianza 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -60400
            TabIndex        =   112
            ToolTipText     =   "Membresia"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1968
         End
         Begin VB.Label lblFiadorLiqCFianza 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -63400
            TabIndex        =   111
            ToolTipText     =   "Membresia"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1848
         End
         Begin VB.Label lblFiadorSalLiquido 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -59440
            TabIndex        =   110
            ToolTipText     =   "Membresia"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1956
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar Trabajo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -63400
            TabIndex        =   109
            Top             =   2160
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label lblFLugarTrabajo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -63400
            TabIndex        =   108
            ToolTipText     =   "Membresia"
            Top             =   2400
            Visible         =   0   'False
            Width           =   3852
         End
         Begin VB.Label lblFiadorIngreso 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -60400
            TabIndex        =   107
            ToolTipText     =   "Membresia"
            Top             =   1200
            Visible         =   0   'False
            Width           =   2892
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso"
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
            Left            =   -60400
            TabIndex        =   106
            Top             =   960
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label lblFiadorNombramiento 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -63400
            TabIndex        =   105
            ToolTipText     =   "Membresia"
            Top             =   1200
            Visible         =   0   'False
            Width           =   2892
         End
         Begin VB.Label Label65 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombramiento"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -63400
            TabIndex        =   104
            Top             =   960
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label lblFiadorEstado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -59440
            TabIndex        =   103
            ToolTipText     =   "Membresia"
            Top             =   600
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.Label lblFiadorMembresia 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -63400
            TabIndex        =   102
            ToolTipText     =   "Membresia"
            Top             =   600
            Visible         =   0   'False
            Width           =   3852
         End
         Begin VB.Label Label62 
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -59440
            TabIndex        =   101
            Top             =   360
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Membresia"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   -63400
            TabIndex        =   100
            Top             =   360
            Visible         =   0   'False
            Width           =   1572
         End
      End
      Begin VB.Frame fraOperacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   612
         Left            =   1080
         TabIndex        =   4
         Top             =   0
         Width           =   12132
         Begin XtremeSuiteControls.PushButton btnConsultaDetalle 
            Height          =   315
            Index           =   0
            Left            =   10200
            TabIndex        =   24
            Top             =   150
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Estudio"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton btnConsultaDetalle 
            Height          =   315
            Index           =   1
            Left            =   11040
            TabIndex        =   25
            Top             =   150
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Trámite"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit txtConsultaId 
            Height          =   312
            Left            =   1440
            TabIndex        =   34
            Top             =   156
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConsultaCedula 
            Height          =   312
            Left            =   3240
            TabIndex        =   35
            Top             =   156
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConsultaNombre 
            Height          =   312
            Left            =   5400
            TabIndex        =   36
            Top             =   156
            Width           =   4812
            _Version        =   1572864
            _ExtentX        =   8488
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption lblOperacion 
            Height          =   615
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12255
            _Version        =   1572864
            _ExtentX        =   21616
            _ExtentY        =   1085
            _StockProps     =   14
            Caption         =   "Operación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption scOperacionBar 
         Height          =   612
         Left            =   0
         TabIndex        =   130
         Top             =   0
         Width           =   12132
         _Version        =   1572864
         _ExtentX        =   21399
         _ExtentY        =   1080
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.96
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.PushButton btnEstudio 
      Height          =   312
      Left            =   10200
      TabIndex        =   22
      Top             =   1560
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Estudio de Crédito"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Index           =   1
      Left            =   6000
      TabIndex        =   14
      Top             =   1080
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   550
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarioValida 
      Height          =   204
      Index           =   0
      Left            =   4440
      TabIndex        =   12
      Top             =   1080
      Width           =   204
      _Version        =   1572864
      _ExtentX        =   360
      _ExtentY        =   360
      _StockProps     =   79
      Appearance      =   16
   End
   Begin FPSpreadADO.fpSpread vGridSolicitudes 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   15015
      _Version        =   524288
      _ExtentX        =   26485
      _ExtentY        =   5318
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
      MaxCols         =   493
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmCR_ComitesAprobaciones.frx":139EF
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtComiteId 
      Height          =   372
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtComiteDesc 
      Height          =   372
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   6492
      _Version        =   1572864
      _ExtentX        =   11451
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Index           =   0
      Left            =   1200
      TabIndex        =   10
      Top             =   1080
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   550
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuarioClave 
      Height          =   312
      Index           =   0
      Left            =   2760
      TabIndex        =   11
      Top             =   1080
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   550
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
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuarioClave 
      Height          =   312
      Index           =   1
      Left            =   7560
      TabIndex        =   15
      Top             =   1080
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   550
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
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarioValida 
      Height          =   204
      Index           =   1
      Left            =   9240
      TabIndex        =   16
      Top             =   1080
      Width           =   204
      _Version        =   1572864
      _ExtentX        =   360
      _ExtentY        =   360
      _StockProps     =   79
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Index           =   2
      Left            =   10800
      TabIndex        =   18
      Top             =   1080
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   550
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuarioClave 
      Height          =   312
      Index           =   2
      Left            =   12360
      TabIndex        =   19
      Top             =   1080
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   550
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
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarioValida 
      Height          =   204
      Index           =   2
      Left            =   14040
      TabIndex        =   20
      Top             =   1080
      Width           =   204
      _Version        =   1572864
      _ExtentX        =   360
      _ExtentY        =   360
      _StockProps     =   79
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboFiltroEstado 
      Height          =   312
      Left            =   12240
      TabIndex        =   21
      Top             =   1560
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2990
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.PushButton btnSolicitud 
      Height          =   312
      Left            =   8160
      TabIndex        =   23
      Top             =   1560
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Solicitud en Trámite"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.PushButton btnActa 
      Height          =   315
      Left            =   3240
      TabIndex        =   27
      Top             =   1560
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   5400
      TabIndex        =   31
      Top             =   1560
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   6720
      TabIndex        =   32
      Top             =   1560
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
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
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":1430F
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":1492B
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":14F49
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":15630
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":15F01
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":16628
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":16C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":174F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":17C05
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":1830C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":18A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":19028
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":19759
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":19E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ComitesAprobaciones.frx":1A55D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnActaInforme 
      Height          =   315
      Left            =   3720
      TabIndex        =   170
      Top             =   1560
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   550
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmCR_ComitesAprobaciones.frx":1AC73
   End
   Begin XtremeSuiteControls.ComboBox cboActas 
      Height          =   330
      Left            =   1200
      TabIndex        =   173
      Top             =   1560
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   3000
      TabIndex        =   180
      Top             =   720
      Width           =   7695
      _Version        =   1572864
      _ExtentX        =   13573
      _ExtentY        =   238
      _StockProps     =   93
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.UpDown UpDownComite 
      Height          =   375
      Left            =   10800
      TabIndex        =   187
      Top             =   240
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   656
      _StockProps     =   64
      Orientation     =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
      BuddyControl    =   ""
      BuddyProperty   =   ""
   End
   Begin VB.Label lblFechas 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   30
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Acta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   26
      Top             =   1560
      Width           =   1092
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario 3:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   9840
      TabIndex        =   17
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario 2:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   5040
      TabIndex        =   13
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comité:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario 1:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1092
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   492
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   14532
      _Version        =   1572864
      _ExtentX        =   25633
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.44
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCR_ComitesAprobaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mId_Comite As Integer, mTipo_Aprobacion As String, vScroll As Boolean
Private mOperacion As String
Private mEstudioCredito As String
Private mNuevoEstado As String
Private mFondoSolidario As Double
Private mLiquidezCFianza As Double
Private mDevengadoMes As Double
Private rslocal As New ADODB.Recordset
Private mFechaInicio As String
Private mCarga As Boolean

Dim vNAprobaciones As Integer, vMancomunado As Boolean, vActa As Long
Dim vRngInicio As Currency, vRngCorte As Currency, vLineaFiltra As Integer

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vPaso As Boolean



Private Sub btnActa_Click()
 If fraActa.Visible Then
    fraActa.Visible = False
 Else
    fraActa.Left = 120
    fraActa.top = 1920
    fraActa.Visible = True
    tcActas.Item(0).Selected = True
 End If
End Sub

Private Sub sbComiteActa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vPaso = True
strSQL = "exec spCrd_Comites_Actas_Abiertas " & txtComiteId.Text
    Call sbCbo_Llena_New(cboActas, strSQL, False, True)
vPaso = False

If cboActas.ListCount >= 0 Then
   Call cboActas_Click
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbActa_Consultar(pActa As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

tcActas.Item(0).Selected = True

lswAsistencia.ListItems.Clear

strSQL = "select * from CRD_COMITES_ACTAS" _
       & " WHERE ID_COMITE = " & txtComiteId.Text & " and ACTA = '" & pActa & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txtActa.Text = rs!acta
    txtSesion.Text = rs!Sesion_Id
    dtpActaFecha.Value = rs!fecha
    txtActaEstado.Text = IIf((rs!Estado = "A"), "Abierta", "Cerrada")
    txtActasNotas.Text = rs!Notas
    vPaso = True
        Call sbCboAsignaDato(cboActas, rs!Sesion_Id, True, rs!acta)
    vPaso = False
Else
    txtActa.Text = ""
    txtSesion.Text = ""
    dtpActaFecha.Value = fxFechaServidor
    txtActaEstado.Text = "Abierta"
    txtActasNotas.Text = ""
End If

rs.Close

'Carga Asistencia
'

vPaso = True

With lswAsistencia
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Identificación", 2100
    .ColumnHeaders.Add , , "Nombre", 4100
    .Checkboxes = True
    
    If Mid(txtActaEstado.Text, 1, 1) = "A" Then
        .Enabled = True
    Else
        .Enabled = False
    End If
    
    If txtActa.Text <> "" Then
            strSQL = "exec spCrd_Comites_Acta_Asistencia_Consulta " & txtComiteId.Text & ",'" & txtActa.Text & "'"
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
             Set itmX = .ListItems.Add(, , rs!Cedula)
                 itmX.SubItems(1) = rs!Nombre
                 If rs!ASISTENCIA = 1 Then
                    itmX.Checked = vbChecked
                    itmX.ForeColor = vbBlue
                 End If
             rs.MoveNext
            Loop
            rs.Close
    End If
    
End With

vPaso = False



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnActaConsulta_Click()
    
On Error GoTo vError
    
Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Comites_Actas_Consulta " & txtComiteId.Text _
       & ", '" & Format(dtpActaInicio.Value, "yyyy-mm-dd") & " 00:00:00'" _
       & ", '" & Format(dtpActaCorte.Value, "yyyy-mm-dd") & " 23:59:59', '" & txtActaFiltro.Text & "'"
Call OpenRecordSet(rs, strSQL)

lswActaH.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswActaH.ListItems.Add(, , rs!id_Comite)
      itmX.SubItems(1) = rs!acta
      itmX.SubItems(2) = rs!Sesion_Id
      itmX.SubItems(3) = Format(rs!fecha, "yyyy-mm-dd")
      itmX.SubItems(4) = rs!ESTADO_DESC
      itmX.SubItems(5) = rs!Comite_Desc
      itmX.SubItems(6) = rs!Registro_Fecha & ""
      itmX.SubItems(7) = rs!Registro_Usuario & ""
      itmX.SubItems(8) = rs!CIERRE_FECHA & ""
      itmX.SubItems(9) = rs!CIERRE_USUARIO & ""
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnActaInforme_Click()
Dim strSubtitulo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
' .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Crédito"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario='Usuario : " & glogon.Usuario & "'"
 .Formulas(2) = "fxFecha='Fecha   : " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 
  strSQL = "{vCrd_Comites_Actas.ID_COMITE} = " & txtComiteId.Text & " AND {vCrd_Comites_Actas.ACTA} = '" & txtActa.Text & "'"
  
    .ReportFileName = SIFGlobal.fxPathReportes("Credito_Comite_Actas.rpt")
    .SelectionFormula = strSQL
    

    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnActaTool_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


Select Case Index
    Case 0 'Nueva

            strSQL = "select Acta, dbo.MyGetdate() as 'Fecha' from COMITES" _
                   & " WHERE ID_COMITE = " & txtComiteId.Text
                   
            strSQL = "exec spCrd_Comites_Acta_Nueva " & txtComiteId.Text & ", '" & glogon.Usuario & "'"
            Call OpenRecordSet(rs, strSQL)
            If Not rs.EOF And Not rs.BOF Then
                txtActa.Text = CStr(rs!acta)
                txtSesion.Text = rs!Sesion
                dtpActaFecha.Value = rs!fecha
                txtActaEstado.Text = "Abierta"
            Else
                txtActa.Text = ""
            End If
            rs.Close
            
            txtActasNotas.Text = ""
            lswAsistencia.ListItems.Clear


    Case 1 'Guardar
    
        If txtActa.Text <> "" Then
                strSQL = "exec spCrd_Comites_Acta " & txtComiteId.Text & ",'" & txtActa.Text & "','" & Format(dtpActaFecha.Value, "yyyy/mm/dd") _
                       & "','" & Trim(txtActasNotas.Text) & "','" & Mid(txtActaEstado.Text, 1, 1) & "','" & glogon.Usuario & "', '" & txtSesion.Text & "'"
                Call ConectionExecute(strSQL)
                
                Call txtComiteId_LostFocus
'                Call sbActa_Consultar(txtActa.Text)
        End If

    Case 2 'Cierra

        If txtActa.Text <> "" Then
                strSQL = "exec spCrd_Comites_Acta_Cierra " & txtComiteId.Text & ",'" & txtActa.Text & "','" & glogon.Usuario & "'"
                Call OpenRecordSet(rs, strSQL)
                
                If rs!Pass = 1 Then
                     MsgBox "Acta Cerrada Satisfactoriamente!", vbInformation
                     Call sbActa_Consultar(txtActa.Text)
                
                Else
                    MsgBox rs!Mensaje, vbExclamation
                End If
        End If

End Select


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCausas_Click()
    Call sbHabilitaControles(True)
    fraCausas.Visible = False
    Call sbLimpiarDatosCreditos
    Call sbCargarListaSolicitudes
End Sub

Private Sub btnConsultaDetalle_Click(Index As Integer)
    Dim frm As Form
    On Error GoTo vError

    If Len(Trim(mOperacion)) = 0 Then
        MsgBox "Debe seleccionar una solicitud"
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Estudio
        
            Dim strSQL As String, rs As New ADODB.Recordset
            Dim x As clsEstudioCrd
            
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
            x.xOperacion = mOperacion
            x.xkey = glogon.ConectRPT
            
            If lblOperacion.Caption = "Operación" Then
            
                strSQL = "select cod_preAnalisis from CRD_PREA_PREANALISIS" _
                           & " Where id_solicitud = " & mOperacion
                
                Call OpenRecordSet(rs, strSQL)
                If rs.EOF And rs.BOF Then
                    Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                                , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                                , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                
                Else
                    x.vSolicitudPreanalisis = rs!cod_PreAnalisis
                    Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                                , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                                , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                End If
                rs.Close
            Else
            
                    x.vSolicitudPreanalisis = mOperacion
                    Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                                , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                                , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
            
            End If
            Set x = Nothing
            
        Case 1  'Tramite
        
            If lblOperacion.Caption = "Operación" Then
                Call sbFormsCall("frmCR_SeguimientoTramites")
                For Each frm In Forms
                    If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                        Call frm.sbConsultaExterna(Val(mOperacion))
                        Exit For
                    End If
                Next frm
            End If

    End Select
    
     Me.MousePointer = vbDefault
    
    Exit Sub
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnEstudio_Click()
        If Len(txtComiteId) = 0 Then
            Exit Sub
        End If
    
        vGridSolicitudes.Col = 3
        vGridSolicitudes.Row = 0
        vGridSolicitudes.Text = "Estudio"
        lblOperacion.Caption = "Estudio"
        
        Call sbLimpiarDatosCreditos
        Call sbCargarListaSolicitudes
End Sub

Private Sub btnExportar_Click(Index As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

If Index = 0 Then
    Call Excel_Exportar_Lsw(lswActaH, ProgressBarX)
Else
    Call Excel_Exportar_Lsw(lswActaR, ProgressBarX)
End If

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnResolucion_Click(Index As Integer)

Dim strSQL As String, rs As New ADODB.Recordset
    
    On Error GoTo vError

    If Len(Trim(mOperacion)) = 0 Then
        MsgBox "Debe seleccionar una solicitud", vbExclamation
        Exit Sub
    End If
    
    txtObservacion.Text = fxSysCleanTxtInject(txtObservacion.Text)
    txtAcuerdoJD.Text = fxSysCleanTxtInject(txtAcuerdoJD.Text)
    
    If Len(Trim(txtObservacion)) <= 50 Then
        MsgBox "Debe indicar una observación válida (50 Caracteres mínimo)", vbExclamation
        Exit Sub
    End If


    Select Case Index
        Case 0 'Aprobar
            Call sbResolucion("A")
        Case 1 'Pendiente
            Call sbResolucion("P")
        Case 2 'Denegar
            Call sbResolucion("D")
            
        Case 3 'VBPENAVALUO
            Call sbResolucion("V")
            
        Case 4 'VBPENLIQ
            Call sbResolucion("VL")
            
    End Select
    
    
    txtObservacion.Text = ""
    txtAcuerdoJD.Text = ""
    
    
    Exit Sub
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnSolicitud_Click()
        If Len(txtComiteId) = 0 Then
            Exit Sub
        End If
    
        vGridSolicitudes.Col = 3
        vGridSolicitudes.Row = 0
        vGridSolicitudes.Text = "Solicitud"
        lblOperacion.Caption = "Operación"
        Call sbLimpiarDatosCreditos
        Call sbCargarListaSolicitudes
End Sub

Private Sub cboActas_Click()
If vPaso Then Exit Sub

On Error GoTo vError

vPaso = True
strSQL = "exec spCrd_Comites_Actas_Abiertas " & txtComiteId.Text
    Call sbCbo_Llena_New(cboActas, strSQL, False, True)
vPaso = False

vGrid.MaxRows = 0
txtActa.Text = ""
txtSesion.Text = ""
txtActaEstado.Text = ""
txtActasNotas.Text = ""

If cboActas.ListCount >= 0 Then
  Call sbActa_Consultar(cboActas.ItemData(cboActas.ListIndex))
End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboFiltroEstado_Click()
    If Not mCarga Then
        Call sbCargarListaSolicitudes
    End If
End Sub

Private Sub sbFechaInicio()
    Dim strSQL As String, rs As New ADODB.Recordset
    
On Error GoTo error
    'Consulta la fecha de inicio de revisiones
    
    Me.MousePointer = vbHourglass
    
    strSQL = "select isnull(valor,'') as valor from CRD_COMITES_PARAMETROS where COD_PARAMETRO ='10'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mFechaInicio = Trim(rs!Valor)
    End If
    rs.Close
   
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
    
End Sub

Private Sub sbCargarListaSolicitudes()
' Carga Lista de operaciones
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Long

On Error GoTo error
    'Consulta la lista de las Operaciones
    
'
'If OptSolicitud.Checked = True Then
'                strSQL = "select dbo.fxSemaforo(R.id_solicitud,R.ID_COMITE,'S') as Semaforo,R.id_solicitud AS cod_preanalisis,R.USERREC AS USUARIO,R.cedula,S.nombre,R.codigo AS cod_linea, R.MONTOSOL AS monto,R.CUOTA as cuota,R.PLAZO AS plazo,R.INT AS tasa,case " & " R.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else R.ESTADOSOL end AS Estado,FECHASOL AS FECHA_CREACION,Garantia, case R.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else R.ESTADOSOL end AS EstadoNuevoPrea,'' AS Acta" & " from reg_creditos R " & " inner join socios S on S.cedula = R.cedula " & " where R.ID_COMITE = " & txtId_Comite.Text
'                If Not IsNothing(mFechaInicio) Then
'                    strSQL = strSQL & " and R.FECHASOL >= '" & mFechaInicio & "'"
'                End If
'
'                strSQL = strSQL & " and R.ESTADOSOL = 'R'"
'
'                strSQL = strSQL & " and dbo.fxCRDTagAprobacion(R.id_solicitud)= 0 "
'                strSQL = strSQL & " order by R.FECHASOL"
'            Else
'                strSQL = "select dbo.fxSemaforo(P.cod_preanalisis,P.ID_COMITE,'S') as Semaforo,P.cod_preanalisis,P.USUARIO,P.cedula,S.nombre,P.cod_linea,P.monto,P.cuota,P.plazo,P.tasa,case " & " P.ESTADO when 'R' then 'Recibido' when 'P' then 'Pendiente' else P.ESTADO end as Estado,FECHA_CREACION,Garantia,CASE P.COD_ESTADO_V2 when 'RECI' then 'Recibido' when 'PRCO' then 'PEN. REV. COMITE' when 'PENVB' then 'VB %LIQ. PEN. AVALUO' when 'AAPA' then 'AVALUO ASIG. PEN. APR'  when 'PEND' then 'Pendiente' else ISNULL(P.COD_ESTADO_V2,'') end as EstadoNuevoPrea,'' AS Acta"
'                strSQL = strSQL & " from crd_prea_preanalisis P " & " inner join socios S on s.cedula = P.cedula "
'                strSQL = strSQL & " where P.tipo_preanalisis = 'E' and P.ID_COMITE = " & txtId_Comite.Text
'
'                If Not IsNothing(mFechaInicio) Then
'                    strSQL = strSQL & " and p.FECHA_CREACION >= '" & mFechaInicio & "'"
'                End If
'
'                strSQL = strSQL & " and p.ESTADO = 'R'"
'
'                strSQL = strSQL & " order by FECHA_CREACION"
'            End If

            
    
    Me.MousePointer = vbHourglass
    
    If lblOperacion.Caption = "Operación" Then
    
            strSQL = "select R.id_solicitud as 'EXPEDIENTE', R.USERREC as 'USUARIO', R.cedula, S.nombre, R.codigo, R.MONTOSOL as 'MONTO', R.CUOTA, R.PLAZO, R.INT AS 'TASA', R.Garantia" _
                    & ", case R.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else R.ESTADOSOL end AS 'ESTADO', FECHASOL AS 'FECHA', Gt.Descripcion as 'Garantia_Desc'" _
                    & ", dbo.fxSemaforo(R.id_solicitud,R.ID_COMITE,'S') as 'Semaforo'" _
                    & " from reg_creditos R " _
                    & " inner join socios S on S.cedula = R.cedula " _
                    & " inner join CRD_COMITES_RNG_GARANTIA G on G.cod_garantia = R.garantia and G.id_comite = R.id_comite" _
                    & " inner join CRD_GARANTIA_TIPOS Gt on R.GARANTIA = Gt.GARANTIA " _
                    & " where R.ID_COMITE = " & txtComiteId.Text _
                    & " and R.MontoSol between G.RNG_INICIO and G.RNG_CORTE"
 
                    '& " and R.MontoSol between " & vRngInicio & " and " & vRngCorte _

'            If mFechaInicio <> Empty Then
'                strSQL = strSQL & " and R.FECHASOL >= '" & mFechaInicio & "'"
'            End If
            
            strSQL = strSQL & " and R.FECHASOL between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
                    
            Select Case cboFiltroEstado.Text
            Case "Todos"
                strSQL = strSQL & " and R.ESTADOSOL in ('P','R')"
            Case "Recibida"
                strSQL = strSQL & " and R.ESTADOSOL = 'R'"
            Case "Pendiente"
                strSQL = strSQL & " and R.ESTADOSOL = 'P'"
            End Select
            
            strSQL = strSQL & " and dbo.fxCRDTagAprobacion(R.id_solicitud)= 0 "
            
            If vLineaFiltra = 1 Then
                strSQL = strSQL & " and R.CODIGO in(SELECT CODIGO FROM CRD_COMITES_LINEAS WHERE ID_COMITE = " & txtComiteId.Text & ")"
            End If
            
    Else
            strSQL = "select P.cod_preanalisis as 'EXPEDIENTE', P.USUARIO, P.cedula, S.nombre, P.cod_linea as 'CODIGO', P.monto, P.cuota, P.plazo, P.tasa, P.Garantia" _
                    & " , case P.ESTADO when 'R' then 'Recibido' when 'P' then 'Pendiente' else P.ESTADO end as 'Estado',FECHA_CREACION AS 'FECHA',Gt.Descripcion as 'Garantia_Desc'" _
                    & " , dbo.fxSemaforo(P.cod_preanalisis,P.ID_COMITE,'P')  as 'Semaforo'" _
                    & " from crd_prea_preanalisis P " _
                    & " inner join socios S on s.cedula = P.cedula " _
                    & " inner join CRD_COMITES_RNG_GARANTIA G on G.cod_garantia = P.garantia and G.id_comite = P.id_comite" _
                    & " inner join CRD_GARANTIA_TIPOS Gt on P.GARANTIA = Gt.GARANTIA " _
                    & " where P.tipo_preanalisis = 'E' and P.ID_COMITE = " & txtComiteId.Text _
                    & " and P.MONTO between G.RNG_INICIO and G.RNG_CORTE"
                    
                    '& "   and P.MONTO  between " & vRngInicio & " and " & vRngCorte
                    
'            If mFechaInicio <> Empty Then
'                strSQL = strSQL & " and p.FECHA_CREACION >= '" & mFechaInicio & "'"
'            End If
        
            strSQL = strSQL & " and P.FECHA_CREACION between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        
            Select Case cboFiltroEstado.Text
            Case "Todos"
                strSQL = strSQL & " and P.ESTADO in ('P','R')"
            Case "Recibida"
                strSQL = strSQL & " and p.ESTADO = 'R'"
            Case "Pendiente"
                strSQL = strSQL & " and P.ESTADO = 'P'"
            End Select
    
            If vLineaFiltra = 1 Then
                strSQL = strSQL & " and P.COD_LINEA in(SELECT CODIGO FROM CRD_COMITES_LINEAS WHERE ID_COMITE = " & txtComiteId.Text & ")"
            End If
    
    
    End If
        
    
'Carga el Grid Principal

With vGridSolicitudes

    .MaxRows = 0
    Call OpenRecordSet(rs, strSQL)
    
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .MaxCols = 14
      
      .Row = .MaxRows
      .Col = 2
        
      Select Case rs!Semaforo
            Case "R" ' Rojo
                .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
            Case "A" ' Amarillo
                .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
            Case "V" ' Verde
                .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
        End Select
      
      .Col = 3
      .Text = rs!Expediente
      .Col = 4
      .Text = rs!Usuario & ""
      .Col = 5
      .Text = RTrim(rs!Cedula & "")
      .Col = 6
      .Text = rs!Nombre & ""
      .Col = 7
      .Text = rs!Codigo & ""
      .Col = 8
      .Text = Format(rs!Monto, "Standard")
      .Col = 9
      .Text = Format(rs!Cuota, "Standard")
      .Col = 10
      .Text = rs!Plazo & ""
      .Col = 11
      .Text = Format(rs!Tasa, "Standard")
      .Col = 12
      .Text = rs!Estado & ""
      .Col = 13
      .Text = rs!fecha & ""
      .Col = 14
      .Text = rs!Garantia_Desc & ""
      
      rs.MoveNext
    Loop
    rs.Close

End With
    
Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub



Private Sub chkUsuarioValida_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

i = Index

    If chkUsuarioValida(i).Value = vbChecked Then
        
        Select Case Index
            Case 0
                If Trim(txtUsuario(0).Text) = Trim(txtUsuario(1).Text) Then
                    chkUsuarioValida(i).Value = vbUnchecked
                    MsgBox "El usuario 1 no puede se igual que el 2, proceda a cambiarlo"
                    Exit Sub
                End If
                
                If Trim(txtUsuario(0).Text) = Trim(txtUsuario(2).Text) Then
                    chkUsuarioValida(i).Value = vbUnchecked
                    MsgBox "El usuario 1 no puede se igual que el 3, proceda a cambiarlo"
                    Exit Sub
                End If
                
            Case 1
            
                If Trim(txtUsuario(0).Text) = Trim(txtUsuario(1).Text) Then
                    chkUsuarioValida(i).Value = vbUnchecked
                    MsgBox "El usuario 1 no puede se igual que el 2, proceda a cambiarlo"
                    Exit Sub
                End If
                
                If Trim(txtUsuario(1).Text) = Trim(txtUsuario(2).Text) Then
                    chkUsuarioValida(i).Value = vbUnchecked
                    MsgBox "El usuario 2 no puede se igual que el 3, proceda a cambiarlo"
                    Exit Sub
                End If
            
            Case 2
        
                If Trim(txtUsuario(0).Text) = Trim(txtUsuario(2).Text) Then
                    chkUsuarioValida(i).Value = vbUnchecked
                    MsgBox "El usuario 1 no puede se igual que el 3, proceda a cambiarlo"
                    Exit Sub
                End If
                
                If Trim(txtUsuario(1).Text) = Trim(txtUsuario(2).Text) Then
                    chkUsuarioValida(i).Value = vbUnchecked
                    MsgBox "El usuario 2 no puede se igual que el 3, proceda a cambiarlo"
                    Exit Sub
                End If
        
        End Select
        
        'Verifica que el usuario sea miembro del comité
        strSQL = "select count(*) as 'Existe'" _
               & " from CRD_COMITES_MIEMBROS Cp inner join CRD_COMITES_AUTORIZADORES Ca on Cp.CEDULA = Ca.CEDULA" _
               & " where Cp.ESTADO = 'A' and Cp.USUARIO = '" & Trim(txtUsuario(i).Text) & "'" _
               & " and Ca.ID_COMITE = " & txtComiteId.Text
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then
            chkUsuarioValida(i).Value = vbUnchecked
            MsgBox "El Usuario: " & Trim(txtUsuario(i).Text) & ", no es miembro de este comité!", vbExclamation
            Exit Sub
        End If
        rs.Close
        
    
        'Verifica Usuario / Cifrado Actual
        strSQL = "exec spSEG_Logon '" & Trim(txtUsuario(i).Text) & "','" & SIFGlobal.fxStringCifrado(Trim(txtUsuarioClave(i).Text)) & "'"
        Call OpenRecordSet(rs, strSQL, 1)
        If Not rs!Existe = 0 Then
            chkUsuarioValida(i).Value = vbChecked
        Else
            chkUsuarioValida(i).Value = vbUnchecked
            MsgBox "Clave de Usuario incorrecta, intente de nuevo"
            txtUsuarioClave(i).SetFocus
        End If
        rs.Close
    End If
    Exit Sub
    
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar_Change(pValor As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 ID_COMITE from comites"
    
    If Len(txtComiteId.Text) > 0 Then
    
        If pValor = 1 Then
           strSQL = strSQL & " where estado = 1 and ID_COMITE > '" & txtComiteId.Text & "' order by ID_COMITE asc"
        Else
           strSQL = strSQL & " where estado = 1 and ID_COMITE < '" & txtComiteId.Text & "' order by ID_COMITE desc"
        End If
        
    Else
        strSQL = strSQL & " where estado = 1 order by ID_COMITE asc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtComiteId.Text = rs!id_Comite
      txtComiteId_LostFocus
    End If
    rs.Close
End If

vScroll = False
pValor = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCargarListaCausas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pTipo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case optCausas.Item(0).Value
    pTipo = "D"
  Case optCausas.Item(1).Value
    pTipo = "P"
End Select



If btnSolicitud.Value = True Then
    mEstudioCredito = fxEstudioCreditoId
Else
    mEstudioCredito = mOperacion
End If


lswCausasList.ListItems.Clear
lswCausasList.ColumnHeaders.Clear
lswCausasList.ColumnHeaders.Add , , "Código", 1200
lswCausasList.ColumnHeaders.Add , , "Descripción", 3200
lswCausasList.ColumnHeaders.Add , , "Fecha", 2800
lswCausasList.ColumnHeaders.Add , , "Usuario", 2800

strSQL = "select Pa.*, Cg.DESCRIPCION " _
       & " from CRD_PREA_GESTION Pa inner join OPERACION_CAUSAS Cg on Pa.COD_CAUSAS = Cg.COD_CAUSAS and Pa.TIPO = Cg.TIPO" _
       & " where Pa.COD_PREANALISIS = '" & mEstudioCredito & "' and Pa.TIPO = '" & pTipo & "'" _
       & " order by REGISTRO_FECHA"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswCausasList.ListItems.Add(, , rs!cod_causas)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Registro_Fecha & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
    
 rs.MoveNext
Loop
rs.Close
 
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()

    vModulo = 3
    
    Me.imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

    dtpActaFecha.Value = fxFechaServidor
    
    dtpActaCorte.Value = dtpActaFecha.Value
    dtpActaInicio.Value = DateAdd("d", -30, dtpActaCorte.Value)
    
    

    dtpCorte.Value = dtpActaFecha.Value
    dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)
    
    
    With lswActaH.ColumnHeaders
        .Clear
        .Add , , "Id Comite", 1000
        .Add , , "Id Acta", 1200, vbCenter
        .Add , , "Sesión", 2800, vbCenter
        .Add , , "Fecha", 2100, vbCenter
        .Add , , "Estado", 1500, vbCenter
        .Add , , "Comité", 3500, vbCenter
        
        .Add , , "Reg.Fecha", 2500, vbCenter
        .Add , , "Reg.Usuario", 2500, vbCenter
        
        .Add , , "Cierra Fecha", 2500, vbCenter
        .Add , , "Cierra Usuario", 2500, vbCenter
    End With

    With lswAsistencia.ColumnHeaders
        .Clear
        .Add , , "Identificación", 2200
        .Add , , "Nombre", 4200
    End With

    vScroll = True
    mCarga = True

    vGridSolicitudes.AppearanceStyle = fxGridStyle
    vGridSeguimiento.AppearanceStyle = fxGridStyle
    
    vGridSolicitudes.MaxRows = 0
    vGridSolicitudes.MaxCols = 14

    
    cboFiltroEstado.Clear
    cboFiltroEstado.AddItem ("Todos")
    cboFiltroEstado.AddItem ("Recibida")
    cboFiltroEstado.AddItem ("Pendiente")
    cboFiltroEstado.Text = "Recibida"
    
    
    mCarga = False
    
    tcMain.Item(0).Selected = True
    tcDetalle.Item(0).Selected = True
    
    mFondoSolidario = fxFondoSolPreanalisis
    
    
    Call sbFechaInicio
    Call sbLimpiarDatos
    

End Sub

Private Sub Form_Resize()
'' Procedimiento para posicionar los controles al max y minimizar la pantalla
On Error GoTo vError
    
    scMain.Width = Me.Width
    imgBanner.Width = Me.Width
     
    If Me.Width > 10000 Then
      
        vGridSolicitudes.Width = Me.Width - 400
        vGridSolicitudes.Height = Me.Height - 2850 - FraControles.Height
        
        cboFiltroEstado.Left = Me.Width - 700 - cboFiltroEstado.Width
        btnEstudio.Left = cboFiltroEstado.Left - btnEstudio.Width - 80
        btnSolicitud.Left = btnEstudio.Left - btnSolicitud.Width - 40
                
        dtpCorte.Left = btnSolicitud.Left - dtpCorte.Width - 120
        dtpInicio.Left = dtpCorte.Left - dtpInicio.Width - 40
        lblFechas.Left = dtpInicio.Left - lblFechas.Width - 40
                
        FraControles.top = vGridSolicitudes.top + vGridSolicitudes.Height
        
        FraControles.Left = vGridSolicitudes.Left
        FraControles.Width = vGridSolicitudes.Width
        
        tcMain.Width = FraControles.Width
        tcDetalle.Width = tcMain.Width
        scOperacionBar.Width = tcMain.Width
        
        vGridSeguimiento.Width = tcMain.Width - 200
        
        lswCausasList.Width = tcMain.Width - (lswCausasList.Left + 600)
        
        txtObservacion.Width = tcMain.Width - (txtObservacion.Left + 600)
        
        fraCausas.Width = Me.Width - 500
        fraCausas.Height = Me.Height - 2300
        
        lsw.Width = fraCausas.Width - 150
        lsw.Height = fraCausas.Height - (lsw.top + 450)
        
'        fraFiadores.Left = (tcMain.Width - fraDetalleCredito.Width) / 2
'        fraDetalleCredito.Left = fraFiadores.Left
        
        vGridFianzas.Width = tcDetalle.Width - (vGridFianzas.Left + 600)
        vGridDeudas.Width = tcDetalle.Width - (vGridDeudas.Left + 600)
        vGridDesembolsos.Width = tcDetalle.Width - (vGridDesembolsos.Left + 600)
        vGridRefundiciones.Width = tcDetalle.Width - (vGridRefundiciones.Left + 600)
        
        
        'Centrado
        fraOperacion.Left = (FraControles.Width - fraOperacion.Width) / 2
        vGrid.Left = (FraControles.Width - vGrid.Width) / 2
                

    End If
    
    Exit Sub

vError:
End Sub




Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
    If btnSolicitud.Value = True Then
        strSQL = "insert operacion_gestion(cod_causas,tipo,id_solicitud,codigo) values('" _
               & Item.Text & "','" & mNuevoEstado & "'," & mOperacion _
               & ",'" & Trim(lblLinea.Tag) & "')"
    Else
        strSQL = "insert CRD_PREA_GESTION(cod_causas,tipo,cod_preanalisis,codigo) values('" _
               & Item.Text & "','" & mNuevoEstado & "','" & mOperacion _
               & "','" & Trim(lblLinea.Tag) & "')"
    End If
Else
    If btnSolicitud.Value = True Then
        strSQL = "delete operacion_gestion where cod_causas = '" & Trim(Item.Text) & "' and tipo = '" _
               & mNuevoEstado & "' and id_solicitud = " & mOperacion
    Else
        strSQL = "delete CRD_PREA_GESTION where cod_causas = '" & Trim(Item.Text) & "' and tipo = '" _
               & mNuevoEstado & "' and cod_preanalisis = '" & mOperacion & "'"
    End If
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:

  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswActaH_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim pComiteId As Long, pActa As String

On Error GoTo vError

pComiteId = Item.Text
pActa = Item.SubItems(1)

If txtComiteId.Text <> pComiteId Then
   txtComiteId.Text = pComiteId
   Call txtComiteId_LostFocus
End If

Call sbActa_Consultar(pActa)

vError:
End Sub

Private Sub lswAsistencia_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError
    
            
    strSQL = "UPDATE CRD_COMITES_ACTAS_ASISTENCIA SET ASISTENCIA = " & IIf(Item.Checked, 1, 0) _
           & ", REGISTRO_FECHA = dbo.myGetdate(), REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
           & " WHERE ID_COMITE = " & txtComiteId.Text & " and ACTA = '" & txtActa.Text & "' and cedula = '" & Item.Text & "'"
    
    Call ConectionExecute(strSQL)
    
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lswFiadores_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Call sbCargaDatosFiadores(Item.Text)
End Sub


Private Sub sbFiadores_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

    On Error GoTo vError
    
    If btnSolicitud.Value = True Then
        strSQL = "SELECT F.CEDULAF AS CEDULA, ISNULL(S.NOMBRE, F.NOMBRE) AS 'NOMBRE' FROM FIADORES F LEFT JOIN SOCIOS S ON F.CEDULAF = S.CEDULA WHERE ID_SOLICITUD =" & mOperacion
    Else
        strSQL = "SELECT P.CEDULA,ISNULL(S.NOMBRE, P.NOMBRE) AS 'NOMBRE'FROM CRD_PREA_PREANALISIS P LEFT JOIN SOCIOS S ON P.CEDULA = S.CEDULA WHERE P.COD_PREANALISIS_REF = '" & mOperacion & "'"
    End If
    
    With lswFiadores.ColumnHeaders
        .Clear
        .Add , , "Identificación", 2100
        .Add , , "Nombre", 3100
    End With

     With lswFiadores.ListItems
        .Clear

        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
               Set itmX = .Add(, , rs!Cedula)
                   itmX.SubItems(1) = rs!Nombre
          rs.MoveNext
        Loop
        rs.Close
     
     End With
     
     
    Exit Sub
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbRefundiciones_Load()
Dim strSQL As String

Dim curCuota As Currency, curMonto As Currency, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

curCuota = 0
curMonto = 0

If btnSolicitud.Value Then
    strSQL = "exec spCrd_Comites_Caso_CRD_Refunde '" & Trim(mOperacion) & "','T'"
Else
    strSQL = "exec spCrd_Comites_Caso_CRD_Refunde '" & Trim(mOperacion) & "','e'"
End If
    
vGridRefundiciones.MaxRows = 0
Call sbCargaGrid(vGridRefundiciones, 13, strSQL)

With vGridRefundiciones

    .MaxRows = .MaxRows - 1
    For i = 1 To .MaxRows
        .Row = i
        .Col = 5
        curMonto = curMonto + CCur(.Text)
        .Col = 6
        curCuota = curCuota + CCur(.Text)
    Next i
    
End With

lblRefundeMonto.Caption = Format(IIf(IsNull(curMonto), 0, curMonto), "Standard")
lblRefundeCuota.Caption = Format(IIf(IsNull(curCuota), 0, curCuota), "Standard")


Me.MousePointer = vbDefault
Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbDesembolsos_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass

    If btnSolicitud.Value = True Then
        
        strSQL = "select isnull(sum(monto),0) as Monto from DESEMBOLSOS " _
            & " where id_solicitud = '" & Trim(mOperacion) & "'"
            
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            lblDesembolsoMonto = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
            lblDesembolsoCuota = Format(0, "Standard")
        End If
        rs.Close
        
        strSQL = "select concepto,monto,0 from DESEMBOLSOS where id_solicitud = '" & Trim(mOperacion) & "'"
    
        vGridDesembolsos.MaxRows = 0
        Call sbCargaGrid(vGridDesembolsos, 3, strSQL)
        vGridDesembolsos.MaxRows = vGridDesembolsos.MaxRows - 1
    Else
        mEstudioCredito = mOperacion
    
        strSQL = "SELECT ISNULL(SUM(MONTO),0) AS MONTO, ISNULL(SUM(CUOTA),0) AS CUOTA FROM CRD_PREA_DETALLE_DESEMBOLSOS " _
            & "WHERE COD_PREANALISIS = '" & Trim(mEstudioCredito) & "'"
            
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            lblDesembolsoMonto = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
            lblDesembolsoCuota = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
        End If
        rs.Close
        
        strSQL = "SELECT DESCRIPCION,MONTO,CUOTA FROM CRD_PREA_DETALLE_DESEMBOLSOS WHERE COD_PREANALISIS = '" & Trim(mEstudioCredito) & "'"
    
        vGridDesembolsos.MaxRows = 0
        Call sbCargaGrid(vGridDesembolsos, 3, strSQL)
        vGridDesembolsos.MaxRows = vGridDesembolsos.MaxRows - 1
    
    End If
        
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub optCausas_Click(Index As Integer)
Call sbCargarListaCausas
End Sub


Private Sub sbActas_Resoluciones()
On Error GoTo vError
    
Me.MousePointer = vbHourglass

    With lswActaR.ColumnHeaders
        .Clear
        .Add , , "Id Comite", 1000
        .Add , , "Id Acta", 1200, vbCenter
        .Add , , "Sesión", 2800, vbCenter
        
        .Add , , "Cedula", 1500, vbCenter
        .Add , , "Nombre", 3500
        .Add , , "Línea", 1200, vbCenter
        .Add , , "Garantía", 2500
        .Add , , "Estado", 2500
        .Add , , "Resolución", 3000
        .Add , , "Notas", 3000
        
        .Add , , "Monto", 2100, vbRightJustify
        .Add , , "Plazo", 1100, vbRightJustify
        .Add , , "Tasa", 1100, vbRightJustify
        .Add , , "Cuota", 1800, vbRightJustify
        
        
        .Add , , "Destino", 3000
        .Add , , "Comite", 3000
        
        .Add , , "Reg.Fecha", 2500, vbCenter
        .Add , , "Reg.Usuario", 2500, vbCenter
        
    End With


strSQL = "select Expediente, Recibe_Usuario, Cedula, Nombre, Cod_Linea, Monto, Cuota, Plazo, [Int]" _
        & "    , Estado, Recibe_Fecha, Garantia, Destino, EstadoResolucion, ResolucionNotas, ResolucionEstadoDesc" _
        & "    , ID_COMITE, ACTA, SESION_ID, ComiteDesc" _
        & " from vCrd_Comites_Actas_Resoluciones R" _
        & " Where id_comite = " & txtComiteId.Text & " And Acta = " & txtActa.Text

Call OpenRecordSet(rs, strSQL)

lswActaR.ListItems.Clear
'Format(rs!fecha, "yyyy-mm-dd")
Do While Not rs.EOF
  Set itmX = lswActaR.ListItems.Add(, , rs!id_Comite)
      itmX.SubItems(1) = rs!acta
      itmX.SubItems(2) = rs!Sesion_Id
      itmX.SubItems(3) = rs!Cedula
      itmX.SubItems(4) = rs!Nombre
      itmX.SubItems(5) = rs!cod_linea
      itmX.SubItems(6) = rs!Garantia
      itmX.SubItems(7) = rs!Estado
      itmX.SubItems(8) = rs!ResolucionEstadoDesc
      itmX.SubItems(9) = rs!ResolucionNotas
      itmX.SubItems(10) = Format(rs!Monto, "Standard")
      itmX.SubItems(11) = rs!Plazo
      itmX.SubItems(12) = Format(rs!Int, "Standard")
      itmX.SubItems(13) = Format(rs!Cuota, "Standard")
      
      itmX.SubItems(14) = rs!Destino & ""
      itmX.SubItems(15) = rs!ComiteDesc
      itmX.SubItems(16) = rs!Recibe_Fecha
      itmX.SubItems(17) = rs!Recibe_Usuario
  
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcActas_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 1 'Actas
      Call btnActaConsulta_Click
      
    Case 2 'Resoluciones
      Call sbActas_Resoluciones
    
End Select
End Sub

Private Sub tcDetalle_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Call Form_Resize
    
    
    Select Case Item.Index
    Case 1
        vGrid.MaxRows = 0
        Call sbClasificacion_Load
        
        vGrid.Visible = False
        
    Case 2
        Call sbPatrimonio_Load
    
    Case 3
        lblDeudasTotal.Caption = Format(0, "Standard")
        lblDeudasCuota.Caption = Format(0, "Standard")
        lblDeducciones.Caption = Format(0, "Standard")
        vGridDeudas.MaxRows = 0
        Call sbDeudas_Load
    Case 4
        lblFianzasMonto.Caption = Format(0, "Standard")
        lblFianzasSaldo.Caption = Format(0, "Standard")
        lblFianzasCuota.Caption = Format(0, "Standard")
        Call sbFianzas_Load
    Case 5
        lblRefundeMonto.Caption = Format(0, "Standard")
        lblRefundeCuota.Caption = Format(0, "Standard")
        Call sbRefundiciones_Load
    Case 6
        lblDesembolsoMonto.Caption = Format(0, "Standard")
        Call sbDesembolsos_Load
    End Select


End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
    Case 1
    
        lblFiadorMembresia.Caption = Empty
        lblFiadorEstado.Caption = Empty
        lblFiadorNombramiento.Caption = Empty
        lblFiadorIngreso.Caption = Empty
        lblFiadorInstitucion.Caption = Empty
        lblFiadorSalLiquido.Caption = Empty
        lblFiadorLiqCFianza.Caption = Empty
        lblFiadorLiqSFianza.Caption = Empty
        lblFLugarTrabajo.Caption = Empty
        lblFiadorLiqCFianzaPorc = Empty
        lblFiadorLiqSFianzaPorc = Empty
                
        Call sbFiadores_Load
        
    Case 5 'Causas
        Call sbCargarListaCausas
        
    End Select
    
Call Form_Resize

End Sub


Private Sub sbResolucion(ByVal vEstado As String)
Dim strSQL As String, rs As New ADODB.Recordset

Dim Cod_Parametro As String, Tag As String, LineaTag As Integer, NotaTag As String

Dim EnviaMensaje As Boolean, Email As String, EmailCC As String
Dim Asunto As String, Cuerpo As String

Dim i As Integer, vMensaje As String
Dim vEstadoComite As String, vEstadoEditable As Integer
    
Dim vIndicadoTrasladoSalario As Boolean, LineaValidar As String
Dim vMontoCreditos As Currency, vMontoMaximo As Currency
Dim vLiquidez As Currency, vLiquidezMinima As Currency, liquidezSinCompAdicional As String
   
Dim pTipo As String
   
On Error GoTo vError
    
    

If btnSolicitud.Value = True Then
    pTipo = "S" 'Solicitiud
Else
    pTipo = "E" 'Estudio de Credito

    strSQL = "exec spCrdPrea_Comite_Validacion_Resolucion '" & mOperacion & "', " & txtComiteId.Text
    Call OpenRecordSet(rs, strSQL)
        
        vLiquidezMinima = rs!LiquidezMinima
        vMontoMaximo = rs!MontoMaximo
        vMontoCreditos = rs!MontoCreditos
        liquidezSinCompAdicional = rs!liquidezSinCompAdicional
        vLiquidez = rs!Liquidez
        vIndicadoTrasladoSalario = IIf((rs!TRASLADA_SALARIO = True Or rs!TRASLADA_SALARIO = 1), True, False)
        LineaValidar = rs!LineaValida
        
    rs.Close
End If
    
    Select Case vEstado
        Case "A"
            Cod_Parametro = "01"
            vEstadoComite = "APRO"
            vEstadoEditable = 0
            
                    If pTipo = "E" Then
                        If LineaValidar = lblLinea.Caption Then
                            If Not vIndicadoTrasladoSalario Then
                                If MsgBox("Se está realizando una aprobación de una línea " & LineaValidar & " sin traslado de salario activo. ¿Desea continuar con la aprobación?", vbYesNo + vbQuestion, "Validaciones Generales") <> vbYes Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                        If vMontoCreditos > vMontoMaximo And vMontoMaximo > 0 Then
                            MsgBox "Excede el monto autorizado de aprobación: " & Format(vMontoMaximo, "Standard") & " para este nivel resolutivo.", vbExclamation
                            Exit Sub
                        End If

                        If vLiquidez < vLiquidezMinima Then
                            MsgBox "No se puede aprobar debido a que no cumple con el % de liquidez mínima (" & vLiquidezMinima & ") requerida para el tipo de garantía, favor validar.", vbExclamation
                            Exit Sub
                        End If

                    End If
        
        Case "P"
            vEstadoComite = "PEND"
            Cod_Parametro = "03"
            vEstadoEditable = 1
        
        Case "D"
            vEstadoComite = "DESC"
            Cod_Parametro = "02"
            vEstadoEditable = 0

'Codigo ASECCSS
'            Dim _frmMotivos As New frmPreaSeleccionarMotivoCambioEstado() '35795 mchaves
'            _frmMotivos.codPreanalisis = mOperacion
'            _frmMotivos.ShowDialog()

        Case "V"
            vEstadoComite = "PENVB"
            Cod_Parametro = "03"
            vEstadoEditable = 1
            vEstado = "P"
        
        Case "VL"
            vEstadoComite = "PNVBL" 'Sol 25929 FAlas
            Cod_Parametro = "03"
            vEstadoEditable = 1
            vEstado = "P"
    End Select
    
    
    '--------------------------------------------------------------------------
    'Valida Seguridad de Pantalla
    
    vMensaje = ""
    For i = 0 To vNAprobaciones - 1
        If chkUsuarioValida.Item(i).Value = xtpUnchecked Then
            vMensaje = vMensaje & " - El Usuario Autorizador No." & i + 1 & ", no ha sido validado!"
        End If
    Next i
    
    strSQL = "select dbo.fxCrd_Comites_Acta_Valida(" & txtComiteId.Text & ",'" & txtActa.Text & "') as 'Resultado'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Resultado = 0 Then
            vMensaje = vMensaje & " - El Acta No." & txtActa.Text & ", no existe o no está abierta!"
    End If
    rs.Close
    
    If Len(vMensaje) > 0 Then
        MsgBox vMensaje, vbExclamation
        Exit Sub
    End If
        
    If mTipo_Aprobacion = "M" Then
        NotaTag = "(" & txtUsuario(0).Text & "," & txtUsuario(1).Text & "," & txtUsuario(2).Text & ") " & txtObservacion
    Else
        NotaTag = "(" & txtUsuario(0).Text & ") " & txtObservacion
    End If
    
    strSQL = "select isnull(valor,'') as valor from CRD_COMITES_PARAMETROS where COD_PARAMETRO ='" & Cod_Parametro & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        Tag = Trim(rs!Valor)
    Else
        MsgBox "No existe en parámetros la información del tag asignado para este movimiento"
        Exit Sub
    End If
    rs.Close
    
    If Tag = Empty Then
        MsgBox "No está definido en parámetros, el tag para este movimiento"
        Exit Sub
    End If
    
    strSQL = "select count(*) from crd_tags where tag_codigo = '" & Tag & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs.Fields(0) = 0 Then
        MsgBox "El tag definido en parámetros para este movimiento, no existe en el catalogo de tags"
        Exit Sub
    End If
    rs.Close
    
    
    
    '-------------------------------------------------> Fin de Validacion
    
    NotaTag = Mid(NotaTag, 1, 998)
    
    ' Insertar el Tag
    If btnSolicitud.Value = True Then
    
        Call sbCrdOperacionTags(Format(mOperacion, "Standard"), Trim(lblLinea.Tag), Tag, "", Trim(NotaTag))

    Else
        
        strSQL = "select isnull(max(linea),0)+1 as Linea from CRD_PREA_TAGS where cod_preanalisis = '" & mOperacion & "'"
   
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            LineaTag = rs!Linea
        End If
        rs.Close
    
        strSQL = "insert CRD_PREA_TAGS (LINEA,CODIGO,COD_PREANALISIS,TAG_CODIGO,ASIGNADO_A,REGISTRO_FECHA,REGISTRO_USUARIO,NOTAS)" _
                 & "values(" & LineaTag _
                 & ",'" & Trim(lblLinea.Tag) _
                 & "'," & mOperacion _
                 & ",'" & Tag _
                 & "','',dbo.MyGetdate(),'" & txtUsuario(0).Text _
                 & "','" & NotaTag & "')"
    
        Call ConectionExecute(strSQL)
    End If
            

'Codigo: ASECCSS
'    strSQL = "UPDATE CRD_COMITES_ACTAS_DETALLE set Estado = 'R', EstadoResolucion='" & EstadoV2 & "',"
'            strSQL = strSQL + " Resolucion='" & Resolucion & "', FechaActualiza=Getdate(), UsuarioActualiza='" & glogon.Usuario & "'"
'            strSQL = strSQL + " WHERE Acta='" & txtActa.Text & "' and CodPreanalisis='" & CodPreanalisis & "'"
'
    strSQL = "exec spCrd_Comites_Resolucion_Add " & txtComiteId.Text & ",'" & txtActa.Text & "','" & txtUsuario(0).Text _
           & "','" & pTipo & "','" & mOperacion & "','" & Mid(txtObservacion.Text, 1, 1000) & "','" & vEstado _
           & "', '" & vEstadoComite & "', " & vEstadoEditable & ", '" & txtAcuerdoJD.Text & "', '" & txtUsuario(1).Text _
           & "', '" & txtUsuario(2).Text & "'"
    Call ConectionExecute(strSQL)

    Call Bitacora("Modifica", IIf((pTipo = "S"), "Solicitud: ", "Estudio de Crédito: ") & mOperacion & " Cambia estado a: " & vEstado)
    
    'Registro de Autorizadores
    strSQL = ""
    For i = 0 To vNAprobaciones - 1
        If chkUsuarioValida.Item(i).Value = xtpChecked Then
            strSQL = strSQL & Space(10) & "exec spCrd_Comites_Resolucion_Autorizadores_Add " & txtComiteId.Text & ",'" & txtActa.Text & "','" & glogon.Usuario _
                   & "','" & pTipo & "','" & mOperacion & "','" & Mid(txtObservacion.Text, 1, 1000) & "','" & vEstado & "','" _
                   & txtUsuario.Item(i).Text & "'"
        End If
    Next i
    If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL)
    End If
       
       
       
    '' Envia mensaje de notificación
    Select Case vEstado
        Case "A"
                strSQL = "select isnull(valor,'') as valor from CRD_COMITES_PARAMETROS where COD_PARAMETRO ='04'"
                Call OpenRecordSet(rs, strSQL)
                If Not rs.EOF Then
                    If Trim(rs!Valor) = "S" Then
                        EnviaMensaje = True
                    Else
                        EnviaMensaje = False
                    End If
                Else
                    EnviaMensaje = False
                End If
                rs.Close
        Case "P"
                strSQL = "select isnull(valor,'') as valor from CRD_COMITES_PARAMETROS where COD_PARAMETRO ='05'"
                Call OpenRecordSet(rs, strSQL)
                If Not rs.EOF Then
                    If Trim(rs!Valor) = "S" Then
                        EnviaMensaje = True
                    Else
                        EnviaMensaje = False
                    End If
                Else
                    EnviaMensaje = False
                End If
                rs.Close
        Case "D"
                strSQL = "select isnull(valor,'') as valor from CRD_COMITES_PARAMETROS where COD_PARAMETRO ='06'"
                Call OpenRecordSet(rs, strSQL)
                If Not rs.EOF Then
                    If Trim(rs!Valor) = "S" Then
                        EnviaMensaje = True
                    Else
                        EnviaMensaje = False
                    End If
                Else
                    EnviaMensaje = False
                    
                    
                End If
                rs.Close
    End Select
    
    If EnviaMensaje = True Then
        Email = Empty
        EmailCC = Empty
        Asunto = Empty
        Cuerpo = Empty
        
        '' Carga el correo del usuario que registro el preanalisis o la solicitud
        If btnSolicitud.Value = True Then
            strSQL = "SELECT isnull(U.EMAIL,'') FROM REG_CREDITOS R INNER JOIN USUARIOS U ON R.USERREC = U.NOMBRE WHERE R.ID_SOLICITUD = " & mOperacion
        Else
            strSQL = "SELECT isnull(U.EMAIL,'') FROM CRD_PREA_PREANALISIS P INNER JOIN USUARIOS U ON P.USUARIO = U.NOMBRE WHERE COD_PREANALISIS = '" & mOperacion & "'"
        End If
        
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            Email = rs.Fields(0)
        Else
            Email = Empty
        End If
        rs.Close
        
        '' Carga los correos que tiene que copiar en el correo
        strSQL = "SELECT ISNULL(VALOR,'') FROM CRD_COMITES_PARAMETROS WHERE COD_PARAMETRO IN ('07','08','09','11','12')"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            If rs.Fields(0) <> Empty Then
                If EmailCC = Empty Then
                    EmailCC = rs.Fields(0)
                Else
                    EmailCC = EmailCC & ";" & rs.Fields(0)
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        
        '' Carga la descripción del tag para ponerlo en el asunto del mensaje
        strSQL = "SELECT DESCRIPCION FROM CRD_TAGS WHERE TAG_CODIGO = '" & Tag & "'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            If btnSolicitud.Value = True Then
                Asunto = rs.Fields(0) & " Solicitud:" & mOperacion
            Else
                Asunto = rs.Fields(0) & " Estudio de Crédito:" & mOperacion
            End If
        Else
            Asunto = mOperacion
        End If
        rs.Close
        
        ''Carga el cuerpo del mensaje
        Cuerpo = NotaTag
        
        strSQL = "exec spSys_CORREO_POOL '" & Trim(Cuerpo) & "','" & Trim(Asunto) & "','P','" & Trim(Email) & "'"
        Call ConectionExecute(strSQL)
        
    End If
    
    '' Carga las Causas y limpia la pantalla si la operacion es aprobada
    If vEstado = "P" Or vEstado = "D" Then
        Call sbCargarCausas(vEstado)
    Else
        Call sbLimpiarDatosCreditos
        Call sbCargarListaSolicitudes
    End If
      
    Exit Sub
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

      
End Sub


Private Sub txtActaFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Id Alterno"
   gBusquedas.Col3Name = "Nombre"
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "Select Cedula,CedulaR, Nombre From Socios"
   gBusquedas.Filtro = ""

   frmBusquedas.Show vbModal
   
   txtActaFiltro.Text = Trim(gBusquedas.Resultado)
End If

End Sub

Private Sub txtComiteId_Change()
    Call sbLimpiarDatos
    Call sbLimpiarDatosCreditos
End Sub

Private Sub sbLimpiarDatos()
Dim i As Integer

On Error GoTo vError

    vGridSolicitudes.MaxRows = 0
  For i = 0 To 2
     txtUsuario(i).Text = ""
     txtUsuarioClave(i).Text = ""
     chkUsuarioValida(i).Value = xtpUnchecked
     
     lblUsuario(i).Visible = False
     txtUsuario(i).Visible = False
     txtUsuarioClave(i).Visible = False
     chkUsuarioValida(i).Visible = False
  Next i
    
    txtComiteDesc.Text = Empty
    mId_Comite = 0
    mTipo_Aprobacion = Empty
    vNAprobaciones = 1
    vMancomunado = False
    vActa = 0
    vRngInicio = 0
    vRngCorte = 0

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbLimpiarDatosCreditos()
On Error GoTo vError

    vGridSeguimiento.MaxRows = 0
    
    txtConsultaId.Text = Empty
    txtConsultaCedula.Text = Empty
    txtConsultaNombre.Text = Empty
    lblMembresia.Caption = Empty
    lblEstadoLaboral.Caption = Empty
    lblEstadoActual.Caption = Empty
    lblMontoApr.Caption = Empty
    lblMonto_Girado.Caption = Empty
    lblMontoRefundicion.Caption = Empty
    lblMontoDesembolsos.Caption = Empty
    lblCuotaRefundicion.Caption = Empty
    lblCuotaDesembolsos.Caption = Empty
    lblTotalCuotas.Caption = Empty
    lblCuota.Caption = Empty
    lblDiferenciaCuota.Caption = Empty
    lblLugarTrabajo.Caption = Empty
        
    lblCA.Caption = Empty
        
    txtObservacion.Text = Empty
    
    
    mOperacion = Empty
    
    tcMain.Item(0).Selected = True
    
    Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub txtComiteId_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtUsuario(0).Visible Then txtUsuario(0).SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Filtro = " and estado = 1"
        gBusquedas.Consulta = "select id_comite, DESCRIPCION " _
                            & " from comites"
        frmBusquedas.Show vbModal
        txtComiteId.Text = gBusquedas.Resultado
        txtComiteDesc.Text = gBusquedas.Resultado2
        If txtUsuario(0).Visible Then txtUsuario(0).SetFocus
        
    End If
End Sub

Private Sub sbCargaDatosComite()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

    
    If Len(Trim(txtComiteId)) = 0 Then Exit Sub
    If Val(txtComiteId) = -1 Then
        txtComiteId.Text = Empty
        Exit Sub
    End If

    mId_Comite = txtComiteId.Text

    strSQL = "select *, dbo.fxCrd_Comites_Acta_Abierta(ID_COMITE) as 'ActaAbierta'" _
         & " from comites where estado = 1 " _
         & " and id_comite = " & mId_Comite
    Call OpenRecordSet(rs, strSQL)
    

 For i = 0 To 2
    lblUsuario.Item(i).Visible = False
    txtUsuario.Item(i).Visible = False
    txtUsuarioClave.Item(i).Visible = False
    chkUsuarioValida.Item(i).Visible = False
 Next i
   
   
If Not rs.EOF And Not rs.BOF Then
    txtComiteDesc.Text = rs!Descripcion
    mTipo_Aprobacion = rs!tipo_aprobacion
    vMancomunado = IIf((rs!tipo_aprobacion = "M"), True, False)
    vNAprobaciones = rs!NAprobaciones
    vActa = rs!acta
    vRngInicio = rs!Rng_Inicio
    vRngCorte = rs!Rng_Corte
    vLineaFiltra = rs!Linea_Filtra
    
    If rs!ActaAbierta <> "-1" Then
        txtActa.Text = rs!ActaAbierta
    Else
        txtActa.Text = ""
    End If
Else
    txtComiteDesc.Text = Empty
    txtComiteId.Text = Empty
    
    vMancomunado = False
    vNAprobaciones = 1
    vActa = 0
    vRngInicio = 0
    vRngCorte = 0
    vLineaFiltra = 0

    txtActa.Text = ""
    txtSesion.Text = ""
    
    
End If

'Activa Autorizadores
For i = 0 To vNAprobaciones - 1
   lblUsuario.Item(i).Visible = True
   txtUsuario.Item(i).Visible = True
   txtUsuarioClave.Item(i).Visible = True
   chkUsuarioValida.Item(i).Visible = True
Next i

rs.Close

vPaso = True
    strSQL = "exec spCrd_Comites_Actas_Abiertas " & txtComiteId.Text
    Call sbCbo_Llena_New(cboActas, strSQL, False, True)
vPaso = False

If cboActas.ListCount >= 0 Then
    Call sbActa_Consultar(cboActas.ItemData(cboActas.ListIndex))
End If

Exit Sub
    
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtComiteId_LostFocus()
    Call sbCargaDatosComite
    
    If Len(Trim(txtComiteId)) > 0 Then
       Call btnSolicitud_Click
    End If
End Sub


Private Sub sbCargaDatosCredito(ByVal Operacion As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

If lblOperacion.Caption = "Operación" Then

strSQL = "exec spCrd_Comites_Caso_CRD '" & mOperacion & "','T'"
Else
strSQL = "exec spCrd_Comites_Caso_CRD '" & mOperacion & "','E'"
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then


    txtConsultaCedula.Text = rs!Cedula
    txtConsultaNombre.Text = rs!Nombre
    lblMembresia.Caption = IIf(IsNull(rs!Membresia), "", rs!Membresia)
    txtConsultaId.Text = rs!Caso_Id
    
    lblLinea.Caption = IIf(IsNull(rs!Codigo), "", rs!Codigo)
    lblLinea.Tag = lblLinea.Caption
    
    lblEstadoLaboral.Caption = rs!EstadoLaboral_Desc
    lblEstadoActual.Caption = rs!EstadoPersona_Desc
    
    lblMontoApr.Caption = Format(IIf(IsNull(rs!Monto), "", rs!Monto), "Standard")
    lblCuota.Caption = Format(IIf(IsNull(rs!Cuota), "", rs!Cuota), "Standard")
    lblMonto_Girado = Format(IIf(IsNull(rs!monto_girado), "", rs!monto_girado), "Standard")
    
    lblMontoDesembolsos.Caption = Format(IIf(IsNull(rs!Desembolso_Monto), "", rs!Desembolso_Monto), "Standard")
    lblCuotaDesembolsos.Caption = Format(IIf(IsNull(rs!DESEMBOLSO_CUOTA), "", rs!DESEMBOLSO_CUOTA), "Standard")
    
    lblMontoRefundicion.Caption = Format(IIf(IsNull(rs!REFUNDE_MONTO), "", rs!REFUNDE_MONTO), "Standard")
    lblCuotaRefundicion.Caption = Format(IIf(IsNull(rs!REFUNDE_CUOTA), "", rs!REFUNDE_CUOTA), "Standard")
    
    lblLugarTrabajo.Caption = IIf(IsNull(rs!LUGAR_TRABAJO), "", rs!LUGAR_TRABAJO)
    
    lblTotalCuotas.Caption = Format(IIf(IsNull(rs!REFUNDE_CUOTA + rs!DESEMBOLSO_CUOTA), "", rs!REFUNDE_CUOTA + rs!DESEMBOLSO_CUOTA), "Standard")
    
    lblDiferenciaCuota.Caption = Format(CCur(lblCuota.Caption) - CCur(lblTotalCuotas.Caption), "Standard")
    
    If CDbl(lblDiferenciaCuota.Caption) > 0 Then
        lblDiferenciaCuota.ForeColor = vbRed
    Else
        lblDiferenciaCuota.ForeColor = vbBlack
    End If

    lblCA.Caption = Format(rs!CA, "Standard")
    
    
    lblClasificacion.Caption = "Clasificación de la Persona: " & rs!COD_CATEGORIA_ASOCIADO & ""
End If
rs.Close

Me.MousePointer = vbDefault

lblDiferenciaCuota.FontBold = True

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbCargaDatosFiadores(ByVal Fiador As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass
    
    If btnSolicitud.Value = True Then
        mEstudioCredito = fxEstudioCreditoId
    Else
        mEstudioCredito = mOperacion
    End If
    mDevengadoMes = 0
    
    If GLOBALES.SysASEVersion Then
       'Modo: ASE
        strSQL = "select dbo.fxEC_Membresia(S.cedula,dbo.MyGetdate()) as Membresia, Est.Descripcion as 'ESTADOACTUAL'" _
                & " ,isnull(S.ESTADOLABORAL,'') as ESTADOLABORAL" _
                & " ,S.FECHAINGRESO, I.DESCRIPCION as INSTITUCION, ISNULL(P.SALARIO_LIQUIDO,0) AS SALARIO_LIQUIDO" _
                & " ,ISNULL(LIQUIDEZ_SIMPLE,0) AS LIQUIDEZ_SIMPLE, ISNULL(LIQUIDEZ_CFIANZAS,0) as LIQUIDEZ_CFIANZAS, UP.DESCRIPCION as 'LUGAR_TRABAJO', P.DEVENGADO_MES " _
                & ", isnull(El.Descripcion,'No Indica') as 'EstadoLaboralDesc'" _
                & " from Socios S inner join INSTITUCIONES I on  S.COD_INSTITUCION = I.COD_INSTITUCION " _
                & " inner join AFI_ESTADOS_PERSONA Est on S.EstadoActual = Est.cod_estado" _
                & "  left join AFI_ESTADO_LABORAL El on S.EstadoLaboral = El.Estado_Laboral" _
                & "  left join UPROGRAMATICA UP on S.UP = UP.CODIGO " _
                & "  left join CRD_PREA_PREANALISIS P on S.cedula = P.cedula and P.COD_PREANALISIS_REF = '" & mEstudioCredito _
                & "' where S.CEDULA = '" & Trim(Fiador) & "'"
    Else
       'Modo: ProGrX
        strSQL = "select dbo.fxEC_Membresia(S.cedula,dbo.MyGetdate()) as Membresia, Est.Descripcion as 'ESTADOACTUAL'" _
                & " ,isnull(S.ESTADOLABORAL,'') as ESTADOLABORAL, Dept.Descripcion as 'LUGAR_TRABAJO'" _
                & " ,S.FECHAINGRESO, I.DESCRIPCION as INSTITUCION, ISNULL(P.SALARIO_LIQUIDO,0) AS SALARIO_LIQUIDO" _
                & " ,ISNULL(LIQUIDEZ_SIMPLE,0) AS LIQUIDEZ_SIMPLE, ISNULL(LIQUIDEZ_CFIANZAS,0) as LIQUIDEZ_CFIANZAS, Dept.DESCRIPCION as 'LUGAR_TRABAJO', P.DEVENGADO_MES " _
                & ", isnull(El.Descripcion,'No Indica') as 'EstadoLaboralDesc'" _
                & " from Socios S inner join INSTITUCIONES I on  S.COD_INSTITUCION = I.COD_INSTITUCION " _
                & " left join AFDepartamentos Dept on S.cod_Institucion = Dept.Cod_Institucion and S.cod_Departamento = Dept.Cod_Departamento" _
                & " inner join AFI_ESTADOS_PERSONA Est on S.EstadoActual = Est.cod_estado" _
                & "  left join AFI_ESTADO_LABORAL El on S.EstadoLaboral = El.Estado_Laboral" _
                & "  left join CRD_PREA_PREANALISIS P on S.cedula = P.cedula and P.COD_PREANALISIS_REF = '" & mEstudioCredito _
                & "' where S.CEDULA = '" & Trim(Fiador) & "'"
    End If
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
    
        lblFiadorMembresia = IIf(IsNull(rs!Membresia), "", rs!Membresia)
        
        lblFiadorNombramiento.Caption = rs!EstadoLaboralDesc
        
        lblFiadorEstado = IIf(IsNull(rs!EstadoActual), "", rs!EstadoActual)
        lblFiadorIngreso = IIf(IsNull(rs!FechaIngreso), "", rs!FechaIngreso)
        lblFiadorInstitucion = IIf(IsNull(rs!Institucion), "", rs!Institucion)
        lblFiadorSalLiquido = Format(IIf(IsNull(rs!SALARIO_LIQUIDO), "", rs!SALARIO_LIQUIDO), "Standard")
        lblFiadorLiqCFianza = Format(IIf(IsNull(rs!LIQUIDEZ_CFIANZAS), "", rs!LIQUIDEZ_CFIANZAS), "Standard")
        lblFiadorLiqSFianza = Format(IIf(IsNull(rs!LIQUIDEZ_SIMPLE), "", rs!LIQUIDEZ_SIMPLE), "Standard")
        
        mDevengadoMes = IIf(IsNull(rs!DEVENGADO_MES), 0, rs!DEVENGADO_MES)
        
        lblFLugarTrabajo = IIf(IsNull(rs!LUGAR_TRABAJO), "", rs!LUGAR_TRABAJO)
        If mDevengadoMes > 0 And lblFiadorLiqSFianza <> Empty Then
            lblFiadorLiqSFianzaPorc = Format(((lblFiadorLiqSFianza.Caption / mDevengadoMes) * 100), "Standard")
        End If
        If mDevengadoMes > 0 And lblFiadorLiqCFianza <> Empty Then
            lblFiadorLiqCFianzaPorc = Format(((lblFiadorLiqCFianza.Caption / mDevengadoMes) * 100), "Standard")
        End If
    
'        lblCA.Caption = Format(rs!CA, "Standard")
    End If
    rs.Close
        
    Me.MousePointer = vbDefault
    Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub sbCargarGridSeguimiento(ByVal Operacion As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass
    If lblOperacion.Caption = "Operación" Then

        strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO from CRD_OPERACION_TAGS OT" _
               & " inner join CRD_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO where OT.ID_SOLICITUD = " & Operacion
    Else
        strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO from CRD_PREA_TAGS OT" _
           & " inner join CRD_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO where OT.COD_PREANALISIS = '" & Operacion & "'"
    End If
            
    vGridSeguimiento.MaxCols = 2
    vGridSeguimiento.MaxRows = 0


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    vGridSeguimiento.MaxRows = vGridSeguimiento.MaxRows + 1
    vGridSeguimiento.Row = vGridSeguimiento.MaxRows
  
    vGridSeguimiento.Col = 1
    vGridSeguimiento.Text = rs!Descripcion
    vGridSeguimiento.TextTip = TextTipFixed
    vGridSeguimiento.TextTipDelay = 1000
    vGridSeguimiento.CellNote = "Usuario: " & rs!Registro_Usuario & "[" & rs!Registro_Fecha & "]"
            
    vGridSeguimiento.Col = 2
    vGridSeguimiento.Value = IIf(IsNull(rs!Notas), "", rs!Notas)
    
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

Private Sub txtComiteId_Validate(Cancel As Boolean)
    If Val(txtComiteId.Text) = 0 Then
        txtComiteId.Text = Empty
    End If
End Sub


Private Sub txtUsuario_Change(Index As Integer)
chkUsuarioValida.Item(Index).Value = vbUnchecked
End Sub

Private Sub txtUsuario_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUsuario(Index).SetFocus
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Filtro = " and estado = 'A'"
        gBusquedas.Consulta = "select nombre, DESCRIPCION " _
                            & " from usuarios"
        frmBusquedas.Show vbModal
        txtUsuario(Index).Text = gBusquedas.Resultado
        txtUsuarioClave(Index).SetFocus
        
    End If

End Sub

Private Sub txtUsuarioClave_Change(Index As Integer)
chkUsuarioValida.Item(Index).Value = vbUnchecked
End Sub


Private Sub UpDownComite_DownClick()
Call FlatScrollBar_Change(0)
End Sub

Private Sub UpDownComite_UpClick()
Call FlatScrollBar_Change(1)
End Sub

Private Sub vGridSolicitudes_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    vGridSolicitudes.Col = 3
    vGridSolicitudes.Row = Row
    
    mOperacion = vGridSolicitudes.Text
    tcMain.Item(0).Selected = True
    
    If Len(Trim(mOperacion)) > 0 Then
    
        tcDetalle.Item(0).Selected = True
        Call sbCargaDatosCredito(mOperacion)
        Call sbCargarGridSeguimiento(mOperacion)
        
    End If
    
End Sub


Public Sub sbCargaGridCheckIni(vGrid As Object, MaxCol As Integer, strSQL As String)
'Procedimiento para cargar grids con el check en la primera columna
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

With vGrid
    .MaxCols = MaxCol + 1
    .MaxRows = 1
    .Row = .MaxRows
    For i = 1 To .MaxCols
     .Col = i
     .Text = ""
    Next i
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      .Row = .MaxRows
      .Col = 1
      
      For i = 2 To .MaxCols
        .Col = i
        .Text = CStr(rs.Fields(i - 2).Value & "")
      Next i
      .MaxRows = .MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End With

Exit Sub

vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbHabilitaControles(ByVal Habilita As Boolean)
On Error GoTo vError

If Habilita = False Then
    
    txtComiteId.Enabled = False
    UpDownComite.Enabled = False
'    txtUsuario(0).Enabled = False
'    txtUsuario2.Enabled = False
'    txtClaveUsuario1.Enabled = False
'    txtClaveUsuario2.Enabled = False
'    chkValidarUsuario1.Enabled = False
'    chkValidarUsuario2.Enabled = False
    
Else

    txtComiteId.Enabled = True
    UpDownComite.Enabled = True
'    txtUsuario(0).Enabled = True
'    txtUsuario2.Enabled = True
'    txtClaveUsuario1.Enabled = True
'    txtClaveUsuario2.Enabled = True
'    chkValidarUsuario1.Enabled = True
'    chkValidarUsuario2.Enabled = True
End If

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargarCausas(ByVal Estado As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

mNuevoEstado = Estado

Call sbHabilitaControles(False)

fraCausas.top = 1560
fraCausas.Left = 120
fraCausas.Visible = True

lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "Causa Id", 1200
    .Add , , "Descripción", 3200
End With


strSQL = "select * from operacion_causas where estado = 1 and tipo = '" & Estado & "'"

vPaso = True

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_causas)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = fxChecked(rs!cod_causas, rs!Tipo)
     
     If itmX.Checked Then itmX.ForeColor = vbBlue
     
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

Private Function fxChecked(vCausa As String, vTipo As String) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

If btnSolicitud.Value = True Then

    strSQL = "select isnull(count(*),0) as Existe from operacion_gestion" _
           & " where cod_causas = '" & vCausa & "' and Tipo = '" & vTipo _
           & "' and id_solicitud = " & mOperacion
           
Else

    strSQL = "select isnull(count(*),0) as Existe from CRD_PREA_GESTION" _
           & " where cod_causas = '" & vCausa & "' and Tipo = '" & vTipo _
           & "' and cod_preanalisis = '" & mOperacion & "'"
           
End If
Call OpenRecordSet(rsX, strSQL, 0)
    fxChecked = IIf((rsX!Existe = 0), False, True)
rsX.Close

End Function


Private Function fxFondoSolPreanalisis() As Double
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

    strSQL = "select isnull(valor,0) from CRD_PREA_PARAMETROS WHERE COD_PARAMETRO = '13'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxFondoSolPreanalisis = rs.Fields(0)
    Else
        fxFondoSolPreanalisis = 0
    End If
    rs.Close

Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

Private Function fxEstudioCreditoId() As String
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

    strSQL = "select isnull(COD_PREANALISIS,0) from CRD_PREA_PREANALISIS WHERE TIPO_PREANALISIS = 'E' and ID_SOLICITUD = " & mOperacion
    
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxEstudioCreditoId = rs.Fields(0)
    Else
        fxEstudioCreditoId = 0
    End If
    rs.Close
    
Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

Private Sub sbClasificacion_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
        
    Me.MousePointer = vbHourglass

    If btnSolicitud.Value = True Then
        mEstudioCredito = fxEstudioCreditoId
    Else
        mEstudioCredito = mOperacion
    End If

    
    If mEstudioCredito = Empty Or mEstudioCredito = 0 Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    strSQL = "select (isnull(LIQUIDEZ_CFIANZAS,0)/isnull(DEVENGADO_MES,1))*100 as 'LiquidezCFianza'" _
           & " from CRD_PREA_PREANALISIS" _
           & " WHERE TIPO_PREANALISIS = 'E' and COD_PREANALISIS = '" & mEstudioCredito & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        mLiquidezCFianza = rs!LiquidezCFianza
    End If
    rs.Close
    
    
    strSQL = "exec spCRDPreaClasificacionNew '" & txtConsultaCedula.Text & "'," & CDbl(mLiquidezCFianza) _
            & ",'" & mEstudioCredito & "'"
    
    Call OpenRecordSet(glogon.Recordset, strSQL)
    Call sbCargaGridLocal(vGrid, 3)

    
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer)

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 0

With glogon.Recordset
    Do While Not .EOF
       vGrid.MaxRows = vGrid.MaxRows + 1
       
       vGrid.Row = vGrid.MaxRows
       vGrid.Col = 1
       vGrid.Text = !Codigo
       
       vGrid.Col = 2
       vGrid.Text = !Descripcion
       
       vGrid.Col = 3
       vGrid.Text = !Razon
    
       vGrid.Col = 1
        Select Case LCase(!Color)
            Case "rojo"
                 vGrid.BackColor = &HFF&
            Case "verde"
                 vGrid.BackColor = &H80FF80
            Case "amarillo"
                vGrid.BackColor = &HFFFF&
        End Select
    
      .MoveNext
    Loop
    .Close
End With


End Sub


Private Sub sbPatrimonio_Load()
Dim strSQL As String, rs As New ADODB.Recordset

txtAhorro.Text = 0
txtAporte.Text = 0
txtCustodia.Text = 0
txtCapitalizacion.Text = 0
txtPatrimonio.Text = 0
txtFondos.Text = 0

lblFechaAhorro.Caption = ""
lblFechaAporte.Caption = ""
lblFechaCustodia.Caption = ""
lblCapitalizado.Caption = ""
  
  
strSQL = "exec spCrd_Comites_Caso_PAT_Integral '" & Trim(txtConsultaCedula.Text) & "'"
Call OpenRecordSet(rs, strSQL)
 
If Not rs.EOF And Not rs.BOF Then
   
   
   txtAhorro.Text = Format(IIf(IsNull(rs!ahorro), 0, rs!ahorro), "Standard")
   txtAporte.Text = Format(IIf(IsNull(rs!Aporte), 0, rs!Aporte), "Standard")
   txtCustodia.Text = Format(IIf(IsNull(rs!Custodia), 0, rs!Custodia), "Standard")
   txtCapitalizacion.Text = Format(IIf(IsNull(rs!capitaliza), 0, rs!capitaliza), "Standard")
      
   txtFondos.Text = Format(IIf(IsNull(rs!FND_AHORROS), 0, rs!FND_AHORROS), "Standard")
   
      
   txtPatrimonio.Text = Format(CCur(txtAhorro.Text) + CCur(txtCustodia.Text) + CCur(txtAporte.Text) + CCur(txtCapitalizacion.Text), "Standard")
   
   lblFechaAhorro.Caption = IIf(IsNull(rs!fecAhorro), "", Format(rs!fecAhorro, "dd/mm/yyyy"))
   lblFechaAporte.Caption = IIf(IsNull(rs!fecaporte), "", Format(rs!fecaporte, "dd/mm/yyyy"))
   lblFechaCustodia.Caption = IIf(IsNull(rs!fecCustodia), "", Format(rs!fecCustodia, "dd/mm/yyyy"))
   lblCapitalizado.Caption = IIf(IsNull(rs!fecCapitaliza), "", Format(rs!fecCapitaliza, "dd/mm/yyyy"))
   
   txtPAT_DisponibleBruto.Text = Format(rs!Pat_Garantia_Total, "Standard")
   txtPAT_Saldos.Text = Format(rs!Pat_Garantia_Saldos, "Standard")
    
   txtPAT_Disponible.Text = Format(rs!Pat_Garantia_Total - rs!Pat_Garantia_Saldos, "Standard")

 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
   
 rs.Close
            
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbDeudas_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass

        mEstudioCredito = fxEstudioCreditoId
        
        strSQL = "exec spCrd_Comites_Caso_Deudas_Rsm '" & txtConsultaCedula _
               & "','" & mEstudioCredito & "'"
            
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            lblDeudasTotal.Caption = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), "Standard")
            lblDeudasCuota.Caption = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
            lblDeducciones.Caption = Format(IIf(IsNull(rs!DEDUCCIONES), 0, rs!DEDUCCIONES), "Standard")
        End If
        rs.Close
        
    Call sbCargaGridDeudas
        
        
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargaGridDeudas()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vMora As Boolean
Dim i As Integer


Me.MousePointer = vbHourglass


vMora = False

With vGridDeudas
 
 .MaxRows = 0
 strSQL = "exec spSys_Consulta_Integrada_Creditos '" & txtConsultaCedula.Text & "','A'"
 
 Call OpenRecordSet(rs, strSQL)

  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    
    For i = 1 To .MaxCols
      .Col = i
      Select Case i
        Case 1 'Status

              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
        
             Select Case rs!ProcesoCod
              Case "N"
       
                If Not IsNull(rs!Referencia) Then
                    If rs!MoraCuota = 0 Then
                       .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
                      .TextTip = TextTipFixed
                      .TextTipDelay = 1000
                      .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
                      .CellNoteIndicatorColor = vbRed
                      .CellNote = "Referencia: " & rs!Referencia
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
              
              .CellNote = "Referencia..:" & rs!Referencia & vbCrLf & "Morosidad:  Cuotas: " & rs!MoraCuota & vbCrLf _
                        & "   Intereses : " & Format(rs!MoraInt, "Standard") & vbCrLf _
                        & "   Cargos    : " & Format(rs!MoraCargos, "Standard") & vbCrLf _
                        & "   Póliza    : " & Format(rs!MoraPoliza, "Standard") & vbCrLf _
                        & "   Principal : " & Format(rs!MoraPrincipal, "Standard") & vbCrLf _
                        & "   Cta.+ Vieja : " & Format(rs!MoraAntigua, "####-##") & vbCrLf _
                        & "   Cta. Ultima : " & Format(rs!MoraUltima, "####-##") & vbCrLf & vbCrLf _
                        & "   Total Mora  : " & Format(rs!MoraInt + rs!MoraCargos + rs!MoraPrincipal + rs!MoraPoliza, "Standard") & vbCrLf _
                        & "   Antiguedad  : " & rs!Antiguedad
            
            End If
        
        Case 2 'Operacion
           .CellTag = CStr(rs!ID_SOLICITUD)
           .Text = CStr(rs!ID_SOLICITUD)

        
        Case 3 'Linea
            .Text = rs!Codigo
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = Trim(rs!LineaX) & vbCrLf & vbCrLf & "Formaliza: " & Format(rs!FechaForp, "dd/mm/yyyy") & vbCrLf _
                       & "Usuario: " & Trim(rs!Userfor) & vbCrLf _
                       & "Oficina: " & rs!OficinaX & vbCrLf & vbCrLf _
                       & "Deductora: " & rs!Deductora & vbCrLf _
                       & "Deduce Planilla: " & rs!ind_deduce_planilla & vbCrLf _
                       & "Factor cálculo: " & rs!Base_Calculo & vbCrLf _
                       & "Divisa: " & rs!Divisa_Desc & vbCrLf & vbCrLf _
                       & "Canal: " & rs!CanalDesc & vbCrLf _
                       & "Actividad: " & rs!ActividadDesc
        
        Case 4 'Monto
            .Text = Format(rs!montoapr, "Standard")
        Case 5 'Saldo
            .Text = Format(rs!Saldo, "Standard")
        Case 6 'Cuota
            .Text = Format(rs!Cuota, "Standard")
        
        Case 7 'Mora Financiera
            .Text = Format(rs!MoraInt + rs!MoraCargos + rs!MoraPrincipal + rs!MoraPoliza, "Standard")
        
        Case 8 'Primer Deduccion
            .Text = Format(rs!PriDeduc, "####-##")
        Case 9 'Ultimo Movimiento
            .Text = Format(rs!FecUlt, "####-##")
        Case 10 'Termina
            .Text = Format((Year(rs!Termina) & Format(Month(rs!Termina), "00")), "####-##")
        
        
        Case 11 'Garantia
            .Text = rs!Garantia
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNote = rs!GarantiaDetalle
        Case 12 'Estado
            .Text = rs!Estado
        Case 13 'Proceso
            .Text = rs!Proceso
        
        Case 14 'Referencia
            .Text = rs!Referencia & ""
        Case 15 'Tasa Original
            .Text = Format(rs!TasaOriginal, "Standard")
        Case 16 'Tasa Actual
            .Text = Format(rs!interesv, "Standard")
        Case 17 'Plazo
            .Text = CStr(rs!Plazo)
      End Select
    Next i
    

    rs.MoveNext
  Loop
  rs.Close
  
End With

  
Me.MousePointer = vbDefault

End Sub


Private Sub sbFianzas_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass

    strSQL = "select isnull(sum(R.montoapr),0) as Monto, isnull(sum(R.cuota),0) as Cuota, isnull(sum(R.saldo),0) as Saldo " _
        & "  from reg_creditos R where R.saldo > 0 and R.estado = 'A' and R.id_solicitud " _
        & " in(select id_solicitud from fiadores where cedulaf = '" & Trim(txtConsultaCedula.Text) & "' and firma = 'S')"
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        lblFianzasMonto = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
        lblFianzasSaldo = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), "Standard")
        lblFianzasCuota = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
    End If
    rs.Close
    
    strSQL = "select R.id_solicitud,R.codigo,dbo.fxCRDNumFiadores(R.id_solicitud) as NFiadores" _
       & ",R.montoapr,R.saldo,R.cuota,S.cedula,S.nombre,isnull(M.intc+M.intm+M.amortiza,0) as MoraMnt ," _
       & " dbo.fxCRDClasificacion(S.cedula,dbo.MyGetdate()) as 'Clasificacion' " _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " left join vista_morosidad M on R.id_solicitud = M.id_solicitud" _
       & " Where R.estado = 'A' and R.id_solicitud in(select id_solicitud from fiadores where cedulaf = '" & Trim(txtConsultaCedula.Text) & "')"
           
    vGridFianzas.MaxRows = 0
    Call sbCargaGrid(vGridFianzas, 10, strSQL)
    vGridFianzas.MaxRows = vGridFianzas.MaxRows - 1
        
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    

    
End Sub


Private Function fxColorCell(ByRef vGrid As Object, _
                             ByVal Row As Integer, _
                             ByVal Col As Integer, _
                             ByVal strcolor As String) As String
vGrid.Row = Row
vGrid.Col = Col
Select Case LCase(strcolor)
    Case "rojo"
         vGrid.BackColor = &HFF&
    Case "verde"
         vGrid.BackColor = &H80FF80
    Case "amarillo"
        vGrid.BackColor = &HFFFF&
End Select
End Function

