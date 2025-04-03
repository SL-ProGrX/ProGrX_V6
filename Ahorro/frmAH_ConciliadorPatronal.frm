VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_ConciliadorPatronal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Conciliación Patronal"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   8175
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   13695
      _Version        =   1441793
      _ExtentX        =   24156
      _ExtentY        =   14420
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
      SelectedItem    =   2
      Item(0).Caption =   "Histórico"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "scMain"
      Item(0).Control(2)=   "gbArchivo"
      Item(1).Caption =   "Conciliación"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "cboLocalizados"
      Item(1).Control(1)=   "Label1(3)"
      Item(1).Control(2)=   "gbNoLocaliza"
      Item(1).Control(3)=   "lswC"
      Item(1).Control(4)=   "btnConciliacion(0)"
      Item(1).Control(5)=   "btnConciliacion(1)"
      Item(1).Control(6)=   "scNoLocalizados"
      Item(2).Caption =   "Resultados"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "lswR"
      Item(2).Control(1)=   "scResult"
      Item(2).Control(2)=   "bntResult(0)"
      Item(2).Control(3)=   "bntResult(1)"
      Item(2).Control(4)=   "cboResult"
      Begin XtremeSuiteControls.ListView lswC 
         Height          =   6735
         Left            =   -70000
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
         _ExtentY        =   11880
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5535
         Left            =   -70000
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   13095
         _Version        =   1441793
         _ExtentX        =   23098
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswR 
         Height          =   5295
         Left            =   0
         TabIndex        =   17
         Top             =   1680
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
         _ExtentY        =   9340
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbNoLocaliza 
         Height          =   6255
         Left            =   -69040
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1441793
         _ExtentX        =   19923
         _ExtentY        =   11033
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ListView lswConSearch 
            Height          =   3975
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   11055
            _Version        =   1441793
            _ExtentX        =   19500
            _ExtentY        =   7011
            _StockProps     =   77
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
            Checkboxes      =   -1  'True
            View            =   3
            FullRowSelect   =   -1  'True
            BackColor       =   16777215
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnConOpcion 
            Height          =   375
            Index           =   0
            Left            =   10080
            TabIndex        =   23
            Top             =   5400
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnConOpcion 
            Height          =   375
            Index           =   1
            Left            =   10560
            TabIndex        =   24
            ToolTipText     =   "Eliiminar Registro"
            Top             =   5400
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":0727
         End
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtIdAlterna 
            Height          =   315
            Left            =   2160
            TabIndex        =   26
            Top             =   960
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   315
            Left            =   4200
            TabIndex        =   27
            Top             =   960
            Width           =   6855
            _Version        =   1441793
            _ExtentX        =   12091
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtConMap 
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   5400
            Width           =   9975
            _Version        =   1441793
            _ExtentX        =   17595
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnConOpcion 
            Height          =   375
            Index           =   2
            Left            =   10800
            TabIndex        =   39
            ToolTipText     =   "Eliiminar Registro"
            Top             =   120
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":0E3D
         End
         Begin XtremeShortcutBar.ShortcutCaption scConCaso 
            Height          =   375
            Left            =   0
            TabIndex        =   31
            Top             =   120
            Width           =   11295
            _Version        =   1441793
            _ExtentX        =   19923
            _ExtentY        =   661
            _StockProps     =   14
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Identificación"
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
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Id Empleado"
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
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   29
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
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
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   28
            Top             =   720
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.GroupBox gbArchivo 
         Height          =   1335
         Left            =   -70000
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   13095
         _Version        =   1441793
         _ExtentX        =   23098
         _ExtentY        =   2355
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Height          =   375
            Left            =   9600
            TabIndex        =   6
            Top             =   120
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":147B
         End
         Begin XtremeSuiteControls.PushButton btnCargar 
            Height          =   375
            Left            =   10080
            TabIndex        =   7
            Top             =   120
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":1B7B
         End
         Begin XtremeSuiteControls.PushButton btnInfo 
            Height          =   375
            Left            =   10560
            TabIndex        =   8
            Top             =   120
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":2294
         End
         Begin XtremeSuiteControls.FlatEdit txtArchivo 
            Height          =   495
            Left            =   2640
            TabIndex        =   9
            Top             =   120
            Width           =   6855
            _Version        =   1441793
            _ExtentX        =   12091
            _ExtentY        =   873
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
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnAplicar 
            Height          =   495
            Left            =   6840
            TabIndex        =   12
            Top             =   720
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            TextAlignment   =   1
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":29AD
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnCancelar 
            Height          =   495
            Left            =   8160
            TabIndex        =   13
            Top             =   720
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Cancelar"
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
            TextAlignment   =   1
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":30D4
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.ComboBox cboTipo 
            Height          =   330
            Left            =   2640
            TabIndex        =   15
            Top             =   720
            Width           =   3735
            _Version        =   1441793
            _ExtentX        =   6588
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
         Begin XtremeSuiteControls.PushButton bntResult 
            Height          =   375
            Index           =   2
            Left            =   11160
            TabIndex        =   41
            ToolTipText     =   "Exportar a Excel"
            Top             =   120
            Width           =   495
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            Picture         =   "frmAH_ConciliadorPatronal.frx":37D4
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Análisis"
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
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo"
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
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   10
            Top             =   120
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.ComboBox cboLocalizados 
         Height          =   330
         Left            =   -67360
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
         _Version        =   1441793
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.PushButton bntResult 
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   32
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_ConciliadorPatronal.frx":40A5
      End
      Begin XtremeSuiteControls.PushButton bntResult 
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   33
         ToolTipText     =   "Exportar a Excel"
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_ConciliadorPatronal.frx":47A5
      End
      Begin XtremeSuiteControls.PushButton btnConciliacion 
         Height          =   375
         Index           =   0
         Left            =   -63520
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_ConciliadorPatronal.frx":5076
      End
      Begin XtremeSuiteControls.PushButton btnConciliacion 
         Height          =   375
         Index           =   1
         Left            =   -63040
         TabIndex        =   36
         ToolTipText     =   "Exportar a Excel"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_ConciliadorPatronal.frx":5776
      End
      Begin XtremeSuiteControls.ComboBox cboResult 
         Height          =   330
         Left            =   480
         TabIndex        =   40
         Top             =   600
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
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
      Begin XtremeShortcutBar.ShortcutCaption scNoLocalizados 
         Height          =   375
         Left            =   -70000
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Indique doble click a la persona que desea conciliar"
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No Localizados"
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
         Height          =   375
         Index           =   3
         Left            =   -69760
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin XtremeShortcutBar.ShortcutCaption scResult 
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   1200
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
         _ExtentY        =   661
         _StockProps     =   14
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
      Begin XtremeShortcutBar.ShortcutCaption scMain 
         Height          =   375
         Left            =   -70000
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   13095
         _Version        =   1441793
         _ExtentX        =   23098
         _ExtentY        =   661
         _StockProps     =   14
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
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   12000
      Top             =   480
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   9480
      TabIndex        =   11
      Top             =   360
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   34
      Top             =   960
      Visible         =   0   'False
      Width           =   13695
      _Version        =   1441793
      _ExtentX        =   24156
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "frmAH_ConciliadorPatronal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean



Private Sub sbCarga_Listado()
Dim rsExcel As New ADODB.Recordset

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "" 'Inicializa Bloque

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
    
  'Cargado
    
With rsExcel
  Do While Not .EOF
    Set itmX = lsw.ListItems.Add(, , !Identificacion)
        itmX.SubItems(1) = !ID_ALTERNA & ""
        itmX.SubItems(2) = !Nombre & ""
        itmX.SubItems(3) = Format(!Patronal, "Standard")
    .MoveNext
  Loop
End With
    
Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    lsw.ListItems.Clear

End Sub


Private Sub bntResult_Click(Index As Integer)
Select Case Index
    Case 0 'Search
     
    Case 1 'Export
      Call sbExport(2)
End Select
End Sub

Private Sub btnAplicar_Click()
    If lsw.ListItems.Count = 0 Then
       MsgBox "No existen casos cargados ...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
End Sub

Private Sub btnBuscar_Click()
txtArchivo.Text = ""

With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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
        
        txtArchivo.Text = .FileName
End With


End Sub

Private Sub btnCancelar_Click()
    lsw.ListItems.Clear
    txtArchivo.Text = ""
End Sub

Private Sub btnCargar_Click()
    Call sbCarga_Listado
End Sub



Private Sub sbExport(Item As Integer)

On Error GoTo vError

Me.MousePointer = vbHourglass

PrgBar.Visible = True

Select Case Item
 Case 0 'Historico
    Call Excel_Exportar_Lsw(lsw, PrgBar)
 Case 1 'Concilia
    Call Excel_Exportar_Lsw(lswC, PrgBar)
 Case 2 'Resultado
    Call Excel_Exportar_Lsw(lswR, PrgBar)

End Select


PrgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnConciliacion_Click(Index As Integer)

Select Case Index
    Case 0 'Search
     
    Case 1 'Export
      Call sbExport(1)
End Select


End Sub

Private Sub btnConOpcion_Click(Index As Integer)


gbNoLocaliza.Visible = False

lswC.Visible = True

End Sub

Private Sub btnInfo_Click()

            

  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: IDENTIFICACION, ID_ALTERNA, NOMBRE, PATRONAL" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"


End Sub




Private Sub Form_Activate()
vModulo = 2

End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboTipo.AddItem "Unicamente Casos Reportados"
cboTipo.ItemData(cboTipo.ListCount - 1) = "R"
cboTipo.AddItem "Análisis Completo"
cboTipo.ItemData(cboTipo.ListCount - 1) = "T"
cboTipo.Text = "Análisis Completo"

cboLocalizados.AddItem "Reportados por el Patrono"
cboLocalizados.ItemData(cboLocalizados.ListCount - 1) = "P"
cboLocalizados.AddItem "Personas en la Base de Datos"
cboLocalizados.ItemData(cboLocalizados.ListCount - 1) = "B"
cboLocalizados.Text = "Reportados por el Patrono"

cboResult.AddItem "Listado Completo"
cboResult.ItemData(cboResult.ListCount - 1) = "C"
cboResult.AddItem "Diferencias"
cboResult.ItemData(cboResult.ListCount - 1) = "D"
cboResult.Text = "Listado Completo"

With lsw.ColumnHeaders
    .Clear
    .Add , , "Identificación", 2100
    .Add , , "Id. Alterna", 2100
    .Add , , "Nombre", 4000
    .Add , , "Aporte Patronal", 2500, vbRightJustify
End With


With lswC.ColumnHeaders
    .Clear
    .Add , , "Identificación", 2100
    .Add , , "Id. Alterna", 2100
    .Add , , "Nombre", 4000
    .Add , , "Aporte Patronal", 2500, vbRightJustify
End With


With lswR.ColumnHeaders
    .Clear
    .Add , , "Identificación", 2100
    .Add , , "Id. Alterna", 2100
    .Add , , "Nombre", 4000
    .Add , , "Aporte Patronal", 2500, vbRightJustify
    .Add , , "Aporte Registrado", 2500, vbRightJustify
    .Add , , "Diferencia", 2500, vbRightJustify
End With

With lswConSearch.ColumnHeaders
    .Clear
    .Add , , "Identificación", 2100
    .Add , , "Id. Alterna", 2100
    .Add , , "Nombre", 4000
    .Add , , "Aporte Patronal", 2500, vbRightJustify
End With


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbProcesar()
Dim lng As Long


On Error GoTo vError

With lsw.ListItems

    For lng = 1 To .Count

    strSQL = strSQL & Space(10) & "exec spPAT_Concilia_Patronal_Registro '" & .Item(lng).Text & "','" & .Item(lng).SubItems(1) _
            & "','" & .Item(lng).SubItems(2) & "', " & CCur(.Item(lng).SubItems(3)) _
            & ", " & cbo.ItemData(cbo.ListIndex) & ", '" & Format(dtpCorte.Value, "yyyy-MM-dd") _
            & "', '" & glogon.Usuario & "', 'A'"
       
       If Len(strSQL) > 20000 Then
          Call ConectionExecute(strSQL)
          If Not glogon.error Then
              strSQL = ""
          End If
       End If

    Next lng

End With

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   If Not glogon.error Then
       strSQL = ""
   End If
End If

'Call Bitacora("Aplica", "Cambio Masivo de " & cboTipo.Text & ", Listado de Excel: Líneas(" & vGrid.MaxRows & ")")

Me.MousePointer = vbDefault



MsgBox "Información Actualizada Satisfactoriamente!", vbInformation

Call sbLimpia

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
    Call sbLimpia


End Sub

Private Sub Form_Resize()

On Error Resume Next

imgBanner.Width = Me.Width

tcMain.Width = Me.Width - 150
tcMain.Height = Me.Height - (tcMain.top + 450)

PrgBar.Width = tcMain.Width

gbArchivo.Width = tcMain.Width

lsw.Width = tcMain.Width
lswC.Width = tcMain.Width
lswR.Width = tcMain.Width

scMain.Width = tcMain.Width
scNoLocalizados.Width = tcMain.Width
scResult.Width = tcMain.Width

lsw.Height = tcMain.Height - (lsw.top + 100)
lswR.Height = tcMain.Height - (lswR.top + 100)
lswC.Height = tcMain.Height - (lswC.top + 100)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lswC_DblClick()
If vPaso Then Exit Sub
If lswC.ListItems.Count = 0 Then Exit Sub

lswC.Visible = False


gbNoLocaliza.Visible = True

txtCedula.Text = lswC.SelectedItem
txtIdAlterna.Text = lswC.SelectedItem.SubItems(1)
txtNombre.Text = lswC.SelectedItem.SubItems(2)

scConCaso.Caption = lswC.SelectedItem + " ¦  " + lswC.SelectedItem.SubItems(1) + " ¦ " + lswC.SelectedItem.SubItems(2)




End Sub

Private Sub lswr_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswR.SortKey = ColumnHeader.Index - 1
  If lswR.SortOrder = 0 Then lswR.SortOrder = 1 Else lswR.SortOrder = 0
  lswR.Sorted = True
End Sub

Private Sub lswC_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswC.SortKey = ColumnHeader.Index - 1
  If lswC.SortOrder = 0 Then lswC.SortOrder = 1 Else lswC.SortOrder = 0
  lswC.Sorted = True
End Sub


Private Sub sbLimpia()

tcMain.Item(0).Selected = True

txtArchivo.Text = ""
lsw.ListItems.Clear

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Select Case Item.Index

    Case 0 'Historico
    
    Case 1 'Conciliacion
    
    
    Case 2 'Resultados
    
End Select

Exit Sub

vError:


End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

strSQL = "select COD_INSTITUCION as 'Idx',  '[' + COD_DIVISA + ']  ' + DESCRIPCION as 'ItmX'" _
       & "  from INSTITUCIONES where ACTIVA = 1" _
       & "  order by COD_INSTITUCION"
Call sbCbo_Llena_New(cbo, strSQL, True, True)

dtpCorte.Value = fxFechaServidor

End Sub


