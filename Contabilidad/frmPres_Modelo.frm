VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPres_Modelo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Modelo Presupuestario"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   9975
      _Version        =   1572864
      _ExtentX        =   17595
      _ExtentY        =   11245
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
      ItemCount       =   4
      Item(0).Caption =   "General"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "btnCatalogo"
      Item(0).Control(1)=   "txtNotas"
      Item(0).Control(2)=   "btnPresupuesto"
      Item(0).Control(3)=   "btnPlanning"
      Item(0).Control(4)=   "Label1(2)"
      Item(0).Control(5)=   "GroupBox1(0)"
      Item(0).Control(6)=   "GroupBox1(1)"
      Item(0).Control(7)=   "btnMapeoSinCC"
      Item(0).Control(8)=   "gbIniciar"
      Item(1).Caption =   "Usuarios"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswUsuarios"
      Item(2).Caption =   "Ajustes"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswAjustes"
      Item(3).Caption =   "Ajustes vrs Usuarios"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "lswAjusteAsg"
      Item(3).Control(1)=   "lswUsuarioAsg"
      Item(3).Control(2)=   "lblAjuste"
      Item(3).Control(3)=   "lblUsuario"
      Begin XtremeSuiteControls.ListView lswUsuarioAsg 
         Height          =   5535
         Left            =   -64960
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
         _Version        =   1572864
         _ExtentX        =   8493
         _ExtentY        =   9763
         _StockProps     =   77
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
      End
      Begin XtremeSuiteControls.ListView lswAjusteAsg 
         Height          =   5535
         Left            =   -69880
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
         _Version        =   1572864
         _ExtentX        =   8493
         _ExtentY        =   9763
         _StockProps     =   77
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
      End
      Begin XtremeSuiteControls.ListView lswAjustes 
         Height          =   5775
         Left            =   -69880
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1572864
         _ExtentX        =   17171
         _ExtentY        =   10186
         _StockProps     =   77
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
      Begin XtremeSuiteControls.ListView lswUsuarios 
         Height          =   5775
         Left            =   -69880
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1572864
         _ExtentX        =   17171
         _ExtentY        =   10186
         _StockProps     =   77
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
      End
      Begin XtremeSuiteControls.GroupBox gbIniciar 
         Height          =   855
         Left            =   480
         TabIndex        =   37
         Top             =   5280
         Width           =   8895
         _Version        =   1572864
         _ExtentX        =   15690
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkReinicio 
            Height          =   255
            Left            =   4920
            TabIndex        =   42
            Top             =   360
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Reinicio Total"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnPresupuestoClean 
            Height          =   615
            Left            =   7080
            TabIndex        =   38
            Top             =   240
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Reiniciar el Presupuesto"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmPres_Modelo.frx":0000
         End
         Begin XtremeSuiteControls.DateTimePicker dtpReinicio 
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   40
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
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
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtpReinicio 
            Height          =   375
            Index           =   1
            Left            =   3360
            TabIndex        =   41
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
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
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Rango de Reinicio"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1935
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   8895
         _Version        =   1572864
         _ExtentX        =   15684
         _ExtentY        =   3408
         _StockProps     =   79
         Caption         =   "Resolución"
         ForeColor       =   4210752
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtResolucionUsuario 
            Height          =   312
            Left            =   1680
            TabIndex        =   23
            Top             =   360
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtResolucionFecha 
            Height          =   312
            Left            =   5280
            TabIndex        =   25
            Top             =   360
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkMostrarInicio 
            Height          =   252
            Left            =   4080
            TabIndex        =   26
            Top             =   720
            Width           =   3372
            _Version        =   1572864
            _ExtentX        =   5948
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Mostrar Dashboard al Inicio?  "
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   312
            Left            =   1680
            TabIndex        =   27
            Top             =   720
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtResolucionNotas 
            Height          =   672
            Left            =   1680
            TabIndex        =   30
            Top             =   1080
            Width           =   7212
            _Version        =   1572864
            _ExtentX        =   12721
            _ExtentY        =   1185
            _StockProps     =   77
            ForeColor       =   0
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
            Caption         =   "Notas de Resolución"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   732
            Index           =   9
            Left            =   480
            TabIndex        =   29
            Top             =   1080
            Width           =   1332
         End
         Begin VB.Label Label1 
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
            Height          =   252
            Index           =   1
            Left            =   480
            TabIndex        =   28
            Top             =   720
            Width           =   1332
         End
         Begin VB.Label Label1 
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
            Height          =   252
            Index           =   8
            Left            =   4080
            TabIndex        =   24
            Top             =   360
            Width           =   1332
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario:"
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
            Index           =   7
            Left            =   480
            TabIndex        =   22
            Top             =   360
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.PushButton btnCatalogo 
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   4560
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Importar Cuentas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmPres_Modelo.frx":057A
      End
      Begin XtremeSuiteControls.PushButton btnPresupuesto 
         Height          =   615
         Left            =   2160
         TabIndex        =   10
         Top             =   4560
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Confección por Partida"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmPres_Modelo.frx":0D69
      End
      Begin XtremeSuiteControls.PushButton btnPlanning 
         Height          =   615
         Left            =   7560
         TabIndex        =   11
         Top             =   4560
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Planificador"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmPres_Modelo.frx":1430
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   795
         Left            =   2160
         TabIndex        =   19
         Top             =   600
         Width           =   7215
         _Version        =   1572864
         _ExtentX        =   12721
         _ExtentY        =   1397
         _StockProps     =   77
         ForeColor       =   0
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   855
         Index           =   1
         Left            =   600
         TabIndex        =   21
         Top             =   3600
         Width           =   8775
         _Version        =   1572864
         _ExtentX        =   15473
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Registro"
         ForeColor       =   4210752
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtRegistroUsuario 
            Height          =   312
            Left            =   1560
            TabIndex        =   33
            Top             =   240
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRegistroFecha 
            Height          =   312
            Left            =   5160
            TabIndex        =   34
            Top             =   240
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
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
            Height          =   252
            Index           =   5
            Left            =   3960
            TabIndex        =   32
            Top             =   240
            Width           =   1332
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario:"
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
            Index           =   4
            Left            =   360
            TabIndex        =   31
            Top             =   240
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.PushButton btnMapeoSinCC 
         Height          =   615
         Left            =   3960
         TabIndex        =   35
         Top             =   4560
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Mapeo Cuentas sin Centro de Costo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmPres_Modelo.frx":1B18
      End
      Begin VB.Label lblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "[...]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69880
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label lblAjuste 
         BackStyle       =   0  'Transparent
         Caption         =   "[...]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -64960
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9000
      TabIndex        =   0
      Top             =   1440
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   720
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
            Picture         =   "frmPres_Modelo.frx":21DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPres_Modelo.frx":5671
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPres_Modelo.frx":8B03
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPres_Modelo.frx":8C21
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodoFiscal 
      Height          =   312
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11880
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   5412
      _Version        =   1572864
      _ExtentX        =   9546
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.ComboBox cboContabilidad 
      Height          =   312
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11880
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
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
      Index           =   10
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   14
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo Fiscal"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1332
   End
End
Attribute VB_Name = "frmPres_Modelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem


Private Sub btnCatalogo_Click()
   Call sbFormsCall("frmPres_Modelo_Cuentas", , , , False, Me)
End Sub

Private Sub btnMapeoSinCC_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

strSQL = "exec spPres_MapeaCuentasSinCentroCosto '" & txtCodigo.Text _
        & "'," & cboContabilidad.ItemData(cboContabilidad.ListIndex) & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Revisión de Mapeo de Cuentas sin Centro de Costo, realizado satisfactoriamente!", vbInformation

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnPlanning_Click()
   Call sbFormsCall("frmPres_Planning", , , , False, Me)
End Sub

Private Sub btnPresupuesto_Click()
   Call sbFormsCall("frmPres_Definicion", vbModal, , , False, Me)
End Sub

Private Sub btnPresupuestoClean_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

'strSQL = "exec spPres_MapeaCuentasSinCentroCosto '" & pModelo & "'," & pContabilidad & ",'" & glogon.Usuario & "'"
'Call ConectionExecute(strSQL)


i = MsgBox("Esta seguro que desea Reiniciar el Presupuesto?", vbYesNo)
If i = vbNo Then Exit Sub

strSQL = "Delete PRES_PRESUPUESTO where COD_MODELO = '" & txtCodigo.Text & "'"
Call ConectionExecute(strSQL)

MsgBox "Modelo de Presupuesto inicializado, vuelva a cargar las cuentas!", vbInformation

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboContabilidad_Click()
If vPaso Then Exit Sub


vPaso = True

 strSQL = "select ID_CIERRE as 'IdX',DESCRIPCION as 'ItmX'" _
        & " From CNTX_CIERRES Where COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
        & " order by INICIO_ANIO desc"
 Call sbCbo_Llena_New(cboPeriodoFiscal, strSQL, False, True)

vPaso = False


End Sub

Private Sub cboPeriodoFiscal_Click()
If vPaso Then Exit Sub

 Call sbLimpiaPantalla
End Sub

Private Sub chkReinicio_Click()
If chkReinicio.Value = xtpChecked Then
    dtpReinicio(0).Enabled = False
Else
    dtpReinicio(0).Enabled = True
End If

dtpReinicio(1).Enabled = dtpReinicio(0).Enabled

End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 COD_MODELO from PRES_MODELOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_MODELO > '" & txtCodigo.Text & "' and cod_contabilidad = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
              & " order by COD_MODELO asc"
    Else
       strSQL = strSQL & " where COD_MODELO < '" & txtCodigo.Text & "' and cod_contabilidad = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
              & " order by COD_MODELO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_MODELO
      Call txtCodigo_LostFocus
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 12
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 12

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
 dtpReinicio(0).Value = CDate("2024-01-01")
 dtpReinicio(1).Value = CDate("2024-12-31")
 
 chkReinicio.Value = xtpChecked
 

vEdita = False
 
 vPaso = True
    strSQL = "select cod_contabilidad as 'IdX', Nombre as 'ItmX' from CNTX_Contabilidades order by cod_Contabilidad"
    Call sbCbo_Llena_New(cboContabilidad, strSQL, False, True)
 
  lswUsuarios.ColumnHeaders.Add , , "Usuario", 2000
  lswUsuarios.ColumnHeaders.Add , , "Nombre", 3000
  lswUsuarios.ColumnHeaders.Add , , "Reg.Fecha", 2100
  lswUsuarios.ColumnHeaders.Add , , "Reg.Usuario", 1800
  
  
  lswAjustes.ColumnHeaders.Add , , "Código", 1200
  lswAjustes.ColumnHeaders.Add , , "Descripción", 3500
  lswAjustes.ColumnHeaders.Add , , "Reg.Fecha", 2200
  lswAjustes.ColumnHeaders.Add , , "Reg.Usuario", 1800
  
  
  lswAjusteAsg.ColumnHeaders.Add , , "Descripción", 4000
  lswUsuarioAsg.ColumnHeaders.Add , , "Usuario", 4000
 vPaso = False

With cboEstado
    .Clear
    .AddItem "Pendiente"
    .AddItem "Autorizado"
    .AddItem "Descartado"
    .Text = "Pendiente"
End With

 Call cboContabilidad_Click
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub



Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False

vCodigo = ""
txtCodigo = ""
txtDescripcion.Text = ""

lblUsuario.Caption = "[...]"
lblUsuario.Tag = ""

lblAjuste.Caption = "[...]"
lblAjuste.Tag = ""

txtNotas.Text = ""
txtRegistroFecha.Text = ""
txtRegistroUsuario.Text = ""

txtResolucionFecha.Text = ""
txtResolucionUsuario.Text = ""
txtResolucionNotas.Text = ""

cboEstado.Enabled = True
cboEstado.Text = "Pendiente"

btnCatalogo.Enabled = False
btnPlanning.Enabled = False
btnPresupuesto.Enabled = False
btnPresupuestoClean.Enabled = False

End Sub



Private Sub sbLista_Usuarios()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

lswUsuarios.ListItems.Clear

strSQL = "exec spPres_Modelo_Usuarios " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
               & ",'" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswUsuarios.ListItems.Add(, , rs!Usuario)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = rs!REGISTRO_FECHA & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
     itmX.Checked = IIf((rs!Activo = 1), True, False)
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


Private Sub sbLista_Ajustes_Usuarios(Optional pTipo As String = "A", Optional pCodigo As String = "")

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

lswAjusteAsg.Checkboxes = False
lswUsuarioAsg.Checkboxes = False


Select Case pTipo
    Case "I"
        lblUsuario.Caption = "[...]"
        lblUsuario.Tag = ""
        
        lblAjuste.Caption = "[...]"
        lblAjuste.Tag = ""
        
        lswAjusteAsg.ListItems.Clear
        lswUsuarioAsg.ListItems.Clear
        
        strSQL = "exec spPres_Modelo_Usuarios_Autorizados " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                       & ",'" & txtCodigo.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         Set itmX = lswUsuarioAsg.ListItems.Add(, , rs!Nombre)
             itmX.Tag = rs!Usuario
         rs.MoveNext
        Loop
        rs.Close
    
    
        strSQL = "exec spPres_Modelo_Ajustes_Autorizados " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                       & ",'" & txtCodigo.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         Set itmX = lswAjusteAsg.ListItems.Add(, , rs!Descripcion)
             itmX.Tag = rs!cod_Ajuste
         rs.MoveNext
        Loop
        rs.Close
    
    Case "A" 'Ajustes vinculados con el Usuario (pCodigo = Usuario)
        lblAjuste.Caption = "[...]"
        lblAjuste.Tag = ""
        
        lswAjusteAsg.ListItems.Clear
        lswAjusteAsg.Checkboxes = True
    
        strSQL = "exec spPres_Modelo_AjUs_Ajustes " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                       & ",'" & txtCodigo.Text & "','" & pCodigo & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         Set itmX = lswAjusteAsg.ListItems.Add(, , rs!Descripcion)
             itmX.Tag = rs!cod_Ajuste
             itmX.Checked = IIf((rs!Activo = 1), True, False)
         rs.MoveNext
        Loop
        rs.Close
    
    Case "U" 'Usuario vinculados con los Ajustes (pCodigo = Ajuste)
        lblUsuario.Caption = "[...]"
        lblUsuario.Tag = ""
        
        lswUsuarioAsg.ListItems.Clear
        lswUsuarioAsg.Checkboxes = True
        
        strSQL = "exec spPres_Modelo_AjUs_Usuarios " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                       & ",'" & txtCodigo.Text & "','" & pCodigo & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         Set itmX = lswUsuarioAsg.ListItems.Add(, , rs!Nombre)
             itmX.Tag = rs!Usuario
             itmX.Checked = IIf((rs!Activo = 1), True, False)
         rs.MoveNext
        Loop
        rs.Close
End Select


vPaso = False


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub







Private Sub lswAjusteAsg_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spPres_Modelo_AjUs_Registro " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                     & ",'" & txtCodigo.Text & "','" & Item.Tag & "','" & lblUsuario.Tag _
                     & "','" & glogon.Usuario & "'," & IIf(Item.Checked, 1, 0)
Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub lswAjusteAsg_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

lblAjuste.Caption = Item.Text
lblAjuste.Tag = Item.Tag

Call sbLista_Ajustes_Usuarios("U", Item.Tag)

End Sub


Private Sub lswAjustes_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spPres_Modelo_Ajustes_Registro " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                     & ",'" & txtCodigo.Text & "','" & Item.Text & "','" & glogon.Usuario & "'," & IIf(Item.Checked, 1, 0)
Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswUsuarioAsg_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spPres_Modelo_AjUs_Registro " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                     & ",'" & txtCodigo.Text & "','" & lblAjuste.Tag & "','" & Item.Tag _
                     & "','" & glogon.Usuario & "'," & IIf(Item.Checked, 1, 0)
Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswUsuarioAsg_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

lblUsuario.Caption = Item.Text
lblUsuario.Tag = Item.Tag

Call sbLista_Ajustes_Usuarios("A", Item.Tag)

End Sub


Private Sub lswUsuarios_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spPres_Modelo_Usuarios_Registro " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
                     & ",'" & txtCodigo.Text & "','" & Item.Text & "','" & glogon.Usuario & "'," & IIf(Item.Checked, 1, 0)
Call ConectionExecute(strSQL)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbLista_Ajustes()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

lswAjustes.ListItems.Clear

strSQL = "exec spPres_Modelo_Ajustes " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
               & ",'" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswAjustes.ListItems.Add(, , rs!cod_Ajuste)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!REGISTRO_FECHA & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
     itmX.Checked = IIf((rs!Activo = 1), True, False)
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


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 1 'Usuario
        Call sbLista_Usuarios
    Case 2 'Ajustes
        Call sbLista_Ajustes
    Case 3 'Usuarios vrs Ajustes
        Call sbLista_Ajustes_Usuarios("I", "")
End Select
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "InsertAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
      Call sbBorrar
    
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "Descripcion"
       gBusquedas.Consulta = "select COD_MODELO,Descripcion from PRES_MODELOS"
       gBusquedas.Filtro = " and cod_contabilidad = " & cboContabilidad.ItemData(cboContabilidad.ListIndex)
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(pCodigo As String)


On Error GoTo vError

If Not fxSIFValidaCadena(pCodigo) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spPres_ModelosConsulta '" & pCodigo & "'" _
                & "," & cboContabilidad.ItemData(cboContabilidad.ListIndex)
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!COD_MODELO
  txtCodigo.Text = rs!COD_MODELO
  
  txtDescripcion = rs!Descripcion & ""
  txtNotas.Text = rs!Notas
   
  vPaso = True
  Call sbCboAsignaDato(cboPeriodoFiscal, rs!Periodo, False)
  Call sbCboAsignaDato(cboEstado, rs!Estado_Desc, False)
  vPaso = False
  
  txtResolucionFecha.Text = rs!Resolucion_Fecha & ""
  txtResolucionUsuario.Text = rs!Resolucion_Usuario & ""
  txtResolucionNotas.Text = rs!Resolucion_Notas & ""
  
  txtRegistroFecha.Text = rs!REGISTRO_FECHA & ""
  txtRegistroUsuario.Text = rs!Registro_Usuario & ""
  
  chkMostrarInicio.Value = rs!Mostrar_Inicio
  
  If Mid(cboEstado.Text, 1, 1) <> "P" Then
     cboEstado.Enabled = False
  Else
     cboEstado.Enabled = True
  End If
  
    btnCatalogo.Enabled = cboEstado.Enabled
    btnPlanning.Enabled = cboEstado.Enabled
    btnPresupuesto.Enabled = cboEstado.Enabled
    btnPresupuestoClean.Enabled = cboEstado.Enabled

    tcMain.Item(0).Selected = True
    tcMain.Item(1).Enabled = True
    tcMain.Item(2).Enabled = True
    tcMain.Item(3).Enabled = True

Else
  MsgBox "No se encontró registro verifique...", vbInformation
    
    tcMain.Item(0).Selected = True
    tcMain.Item(1).Enabled = False
    tcMain.Item(2).Enabled = False
    tcMain.Item(3).Enabled = False
    
    
    btnCatalogo.Enabled = False
    btnPlanning.Enabled = False
    btnPresupuesto.Enabled = False
    btnPresupuestoClean.Enabled = False
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtCodigo.Text = "" Then vMensaje = vMensaje & " - Codigo de Modelo no Valido ..." & vbCrLf
If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion no es válido ..."

If (vEdita = False) And (Mid(cboEstado.Text, 1, 1) <> "P") Then vMensaje = vMensaje & " -El Estado para guardar el Modelo no es valido" & vbCrLf
 
 

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strNotas As String

On Error GoTo vError

If Mid(cboEstado.Text, 1, 1) = "P" Then
    strNotas = txtNotas.Text
Else
    strNotas = txtResolucionNotas.Text
End If

strSQL = "exec spPres_ModelosRegistra '" & Trim(txtCodigo.Text) & "'," & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & "," & cboPeriodoFiscal.ItemData(cboPeriodoFiscal.ListIndex) & ",'" & Trim(txtDescripcion.Text) _
       & "','" & Trim(strNotas) & "','" & Mid(cboEstado.Text, 1, 1) & "','" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Modelo Presupuestario:  " & vCodigo & ", Conta Id:" & cboContabilidad.ItemData(cboContabilidad.ListIndex))


vCodigo = txtCodigo.Text

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(txtCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
'  strSQL = "delete PRES_MODELOS where COD_MODELO = '" & vCodigo & "'"
'  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Modelo Presupuestario:  " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_MODELO"
  gBusquedas.Orden = "COD_MODELO"
  gBusquedas.Consulta = "select COD_MODELO,Descripcion from PRES_MODELOS"
  gBusquedas.Filtro = " and cod_contabilidad = " & cboContabilidad.ItemData(cboContabilidad.ListIndex)
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()

If txtCodigo.Text <> "" Then
  Call sbConsulta(txtCodigo.Text)
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select COD_MODELO,Descripcion from PRES_MODELOS"
  gBusquedas.Filtro = " And cod_contabilidad = " & gCntX_Parametros.CodigoConta

  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

