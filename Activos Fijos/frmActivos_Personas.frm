VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmActivos_Personas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición de responsables"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7575
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   11295
      _Version        =   1441792
      _ExtentX        =   19923
      _ExtentY        =   13361
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
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Persona"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "cbo"
      Item(0).Control(2)=   "cboSec"
      Item(0).Control(3)=   "Label1(0)"
      Item(0).Control(4)=   "Label1(1)"
      Item(1).Caption =   "Nómina"
      Item(1).ControlCount=   12
      Item(1).Control(0)=   "chkMarcarTodos"
      Item(1).Control(1)=   "fraCambio"
      Item(1).Control(2)=   "vGridNomina"
      Item(1).Control(3)=   "cboBuscaDept"
      Item(1).Control(4)=   "cboBuscaSec"
      Item(1).Control(5)=   "txtBuscaCedula"
      Item(1).Control(6)=   "txtBuscaNombre"
      Item(1).Control(7)=   "btnBuscar(0)"
      Item(1).Control(8)=   "btnBuscar(1)"
      Item(1).Control(9)=   "scSubTitulos(0)"
      Item(1).Control(10)=   "btnSinc(2)"
      Item(1).Control(11)=   "btnBuscar(2)"
      Begin XtremeSuiteControls.CheckBox chkMarcarTodos 
         Height          =   216
         Left            =   600
         TabIndex        =   17
         Top             =   960
         Width           =   216
         _Version        =   1441792
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         BackColor       =   -2147483633
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   372
         Index           =   0
         Left            =   8760
         TabIndex        =   16
         Top             =   910
         Width           =   912
         _Version        =   1441792
         _ExtentX        =   1609
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Appearance      =   6
         Picture         =   "frmActivos_Personas.frx":0000
         ImageAlignment  =   0
      End
      Begin VB.Frame fraCambio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
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
         Height          =   3495
         Left            =   1680
         TabIndex        =   7
         Top             =   2400
         Width           =   7095
         Begin XtremeSuiteControls.ComboBox cboCambioDept 
            Height          =   312
            Left            =   1920
            TabIndex        =   8
            Top             =   1800
            Width           =   4932
            _Version        =   1441792
            _ExtentX        =   8705
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboCambioSec 
            Height          =   312
            Left            =   1920
            TabIndex        =   9
            Top             =   2160
            Width           =   4932
            _Version        =   1441792
            _ExtentX        =   8705
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFecha 
            Height          =   312
            Left            =   1920
            TabIndex        =   10
            Top             =   1440
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
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
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   312
            Left            =   480
            TabIndex        =   20
            Top             =   960
            Width           =   1572
            _Version        =   1441792
            _ExtentX        =   2773
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   312
            Left            =   2040
            TabIndex        =   21
            Top             =   960
            Width           =   4692
            _Version        =   1441792
            _ExtentX        =   8276
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnBarra 
            Height          =   372
            Index           =   0
            Left            =   4920
            TabIndex        =   23
            Top             =   2880
            Width           =   1392
            _Version        =   1441792
            _ExtentX        =   2455
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Appearance      =   6
            Picture         =   "frmActivos_Personas.frx":0700
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton btnBarra 
            Height          =   372
            Index           =   1
            Left            =   6360
            TabIndex        =   24
            Top             =   2880
            Width           =   432
            _Version        =   1441792
            _ExtentX        =   762
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   6
            Picture         =   "frmActivos_Personas.frx":0E27
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   252
            Index           =   2
            Left            =   480
            TabIndex        =   28
            Top             =   2160
            Width           =   1812
            _Version        =   1441792
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Sección"
            BackColor       =   -2147483633
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
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   252
            Index           =   1
            Left            =   480
            TabIndex        =   27
            Top             =   1800
            Width           =   1812
            _Version        =   1441792
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Departamento"
            BackColor       =   -2147483633
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
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   252
            Index           =   0
            Left            =   480
            TabIndex        =   26
            Top             =   1440
            Width           =   1812
            _Version        =   1441792
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplicación"
            BackColor       =   -2147483633
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
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   252
            Index           =   10
            Left            =   480
            TabIndex        =   25
            Top             =   720
            Width           =   1812
            _Version        =   1441792
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Persona"
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
         End
         Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
            Height          =   492
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   7092
            _Version        =   1441792
            _ExtentX        =   12509
            _ExtentY        =   868
            _StockProps     =   14
            Caption         =   "Cambio de área de trabajo"
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
            VisualTheme     =   3
            Alignment       =   1
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5655
         Left            =   -69880
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   10935
         _Version        =   524288
         _ExtentX        =   19288
         _ExtentY        =   9975
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmActivos_Personas.frx":153D
         VScrollSpecial  =   -1  'True
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   -67360
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   6492
         _Version        =   1441792
         _ExtentX        =   11456
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboSec 
         Height          =   312
         Left            =   -67360
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   6492
         _Version        =   1441792
         _ExtentX        =   11456
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread vGridNomina 
         Height          =   5655
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   11175
         _Version        =   524288
         _ExtentX        =   19711
         _ExtentY        =   9975
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
         MaxCols         =   8
         SpreadDesigner  =   "frmActivos_Personas.frx":1BCA
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboBuscaDept 
         Height          =   312
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   3132
         _Version        =   1441792
         _ExtentX        =   5530
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboBuscaSec 
         Height          =   312
         Left            =   5280
         TabIndex        =   13
         Top             =   480
         Width           =   3132
         _Version        =   1441792
         _ExtentX        =   5530
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtBuscaCedula 
         Height          =   312
         Left            =   2160
         TabIndex        =   14
         Top             =   960
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBuscaNombre 
         Height          =   312
         Left            =   3720
         TabIndex        =   15
         Top             =   960
         Width           =   4692
         _Version        =   1441792
         _ExtentX        =   8276
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   372
         Index           =   1
         Left            =   9720
         TabIndex        =   18
         ToolTipText     =   "Boleta de Asignación de Activos"
         Top             =   910
         Width           =   432
         _Version        =   1441792
         _ExtentX        =   762
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   6
         Picture         =   "frmActivos_Personas.frx":2732
      End
      Begin XtremeSuiteControls.PushButton btnSinc 
         Height          =   315
         Index           =   2
         Left            =   8760
         TabIndex        =   29
         Top             =   480
         Width           =   1392
         _Version        =   1441792
         _ExtentX        =   2455
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Sinc. RRHH"
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
         Appearance      =   6
         Picture         =   "frmActivos_Personas.frx":2E39
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Index           =   2
         Left            =   10200
         TabIndex        =   30
         ToolTipText     =   "Contrato de Responsabilidad"
         Top             =   910
         Width           =   435
         _Version        =   1441792
         _ExtentX        =   762
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   6
         Picture         =   "frmActivos_Personas.frx":3552
      End
      Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
         Height          =   492
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   840
         Width           =   11292
         _Version        =   1441792
         _ExtentX        =   19918
         _ExtentY        =   868
         _StockProps     =   14
         Caption         =   "Responsable Actual"
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
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sección"
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
         Left            =   -69040
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
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
         Left            =   -69040
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Responsables de Activos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   2
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   5892
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmActivos_Personas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSeccion As String, vPaso As Boolean, vCarga As Boolean


Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vBoleta As String

On Error GoTo vError


Select Case Index
  Case 0 'Aplicar
            
     strSQL = "exec spActivos_DepartamentoCambio '" & txtCedula.Text & "','" & cboCambioDept.ItemData(cboCambioDept.ListIndex) _
            & "','" & cboCambioSec.ItemData(cboCambioSec.ListIndex) & "','" & glogon.Usuario & "','" & Format(dtpFecha.Value, "yyyy/mm/dd") & "'"
     Call OpenRecordSet(rs, strSQL, 0)
        vBoleta = rs!Boleta
     rs.Close
     
     
     Call Bitacora("Aplica", "Cambio de Departamento/Sección a: " & txtCedula.Text & " Boleta.:" & vBoleta)
     Call sbBoletaTraslado(vBoleta)
     
     MsgBox "Cambio realizado Satisfactoriamente!", vbInformation
     
  Case 1 'Cerrar
     fraCambio.Visible = False
End Select

fraCambio.Visible = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnBuscar_Click(Index As Integer)
Dim i As Long

Select Case Index
    Case 0 'Buscar
        Call sbNomina
        
    Case 1  'Imprimir
        With vGridNomina
          For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            
            If .Value = vbChecked Then
               .Col = 4
               Call sbBoletaResponsable(.Text, "B")
            End If
            
          Next i
        End With


    Case 2  'Contrato
        With vGridNomina
          For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            
            If .Value = vbChecked Then
               .Col = 4
               Call sbBoletaResponsable(.Text, "C")
            End If
            
          Next i
        End With

End Select
End Sub

Private Sub btnSinc_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spActivos_Sincroniza_RH"
Call ConectionExecute(strSQL)


MsgBox "Sincronización con RRHH finalizada!", vbInformation


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboBuscaDept_Click()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select rtrim(cod_seccion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_Secciones" _
       & " where cod_departamento = '" & cboBuscaDept.ItemData(cboBuscaDept.ListIndex) & "' order by cod_seccion"
Call sbCbo_Llena_New(cboBuscaSec, strSQL, True, True)

 

End Sub

Private Sub cboCambioDept_Click()
Dim strSQL As String

If vPaso Then Exit Sub



strSQL = "select rtrim(cod_seccion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_secciones where cod_departamento = '" & cboCambioDept.ItemData(cboCambioDept.ListIndex) _
       & "' order by cod_seccion"
Call sbCbo_Llena_New(cboCambioSec, strSQL, False, True)

End Sub


Private Sub cbo_Click()
Dim strSQL As String

If vPaso Then Exit Sub


vPaso = True


strSQL = "select rtrim(cod_seccion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_secciones where cod_departamento = '" & cbo.ItemData(cbo.ListIndex) _
       & "' order by cod_seccion"
Call sbCbo_Llena_New(cboSec, strSQL, False, True)

vPaso = False

Call cboSec_Click

End Sub


Private Sub cboSec_Click()
Dim strSQL As String

If vPaso Then Exit Sub


strSQL = "select * from Activos_Personas" _
      & " where cod_departamento = '" & cbo.ItemData(cbo.ListIndex) _
      & "' and cod_seccion = '" & cboSec.ItemData(cboSec.ListIndex) _
      & "' order by identificacion"
Call sbCargaGridLocal(vGrid, strSQL)

End Sub


Private Sub sbCargaGridLocal(ByRef pGrid As Object, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strUltimaSeleccion As String


Me.MousePointer = vbHourglass

On Error GoTo vError

pGrid.MaxRows = 0
pGrid.MaxRows = 1
pGrid.Row = pGrid.MaxRows


Call OpenRecordSet(rs, strSQL, 0)

With pGrid
Do While Not rs.EOF
  .Row = pGrid.MaxRows
  .Col = 1
  
    For i = 1 To 4
      .Col = i
      Select Case i
       Case 1 'Identificacion
          .Text = rs!Identificacion
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = "Registrado: " & rs!registro_usuario & vbCrLf & "Fecha: " & rs!registro_fecha & vbCrLf & vbCrLf _
                    & "Modificado: " & rs!Modifica_Usuario & vbCrLf & "Fecha: " & rs!Modifica_Fecha
       Case 2 'Nombre
          .Text = rs!Nombre
       Case 3 'Alterno
          .Text = rs!cod_Alterno
       
       Case 4 'Activo
          .Value = rs!activo
      End Select
    Next i
  
  pGrid.MaxRows = pGrid.MaxRows + 1
  
  rs.MoveNext

Loop

End With

rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub chkMarcarTodos_Click()
Dim i As Long

With vGridNomina
  For i = 1 To .MaxRows
     .Row = i
     .Col = 1
     .Value = chkMarcarTodos.Value
  Next i
End With

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


vModulo = 36

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

dtpFecha.Value = gActivos.Periodo

vPaso = True
    strSQL = "select rtrim(cod_departamento) as 'IdX', rtrim(descripcion) as 'ItmX'" _
           & " from Activos_departamentos order by cod_departamento"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     cbo.AddItem rs!itmX
     cboCambioDept.AddItem rs!itmX
     
     cbo.ItemData(cbo.ListCount - 1) = CStr(rs!IdX)
     cboCambioDept.ItemData(cboCambioDept.ListCount - 1) = CStr(rs!IdX)
     
     rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
      rs.MoveFirst
      cbo.Text = rs!itmX
      cboCambioDept.Text = rs!itmX
    End If
    rs.Close
vPaso = False

'Busquedas
vPaso = True
    strSQL = "select rtrim(cod_departamento) as 'IdX' , rtrim(descripcion) as 'itmX'" _
           & " from Activos_departamentos order by cod_departamento"
    Call sbCbo_Llena_New(cboBuscaDept, strSQL, True, True)
    
vPaso = False

Call cbo_Click
Call cboCambioDept_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from Activos_Personas" _
       & " where identificacion = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  vGrid.Col = 1
  strSQL = "insert into Activos_Personas(cod_departamento,cod_seccion,identificacion,nombre,cod_alterno,Activo,registro_usuario,registro_fecha) values('" _
         & cbo.ItemData(cbo.ListIndex) & "','" & cboSec.ItemData(cboSec.ListIndex) & "','" & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',getdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Persona : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Activos_Personas set nombre = '" & vGrid.Text & "',cod_departamento = '" _
        & cbo.ItemData(cbo.ListIndex) & "',cod_seccion = '" & cboSec.ItemData(cboSec.ListIndex) _
        & "',modifica_usuario = '" & glogon.Usuario & "',modifica_fecha = getdate()"
 
 vGrid.Col = 3
 strSQL = strSQL & ",cod_alterno = '" & vGrid.Text & "',activo = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & " where identificacion = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Persona: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbNomina()
Dim strSQL As String, vWhere As Boolean

On Error GoTo vError

vPaso = True
vWhere = False

fraCambio.Visible = False

strSQL = "select " & chkMarcarTodos.Value & ",'','',Per.IDENTIFICACION,  Per.NOMBRE, Dept.DESCRIPCION as 'Departamento', Sec.DESCRIPCION as 'Secciones', Per.ACTIVO " _
       & " from ACTIVOS_PERSONAS Per inner join ACTIVOS_DEPARTAMENTOS Dept on Per.COD_DEPARTAMENTO = Dept.COD_DEPARTAMENTO" _
       & " inner join ACTIVOS_SECCIONES Sec on Per.COD_DEPARTAMENTO = Sec.COD_DEPARTAMENTO and Per.COD_SECCION = Sec.COD_SECCION"

If Len(txtBuscaCedula.Text) > 0 Then
  strSQL = strSQL & IIf(vWhere, " AND ", " WHERE ") _
         & " Per.Identificacion like '%" & txtBuscaCedula.Text & "%'"
  vWhere = True
End If

If Len(txtBuscaNombre.Text) > 0 Then
  strSQL = strSQL & IIf(vWhere, " AND ", " WHERE ") _
         & " Per.Nombre like '%" & txtBuscaNombre.Text & "%'"
  vWhere = True
End If

If cboBuscaDept.Text <> "TODOS" And Len(cboBuscaDept.Text) > 0 Then
  strSQL = strSQL & IIf(vWhere, " AND ", " WHERE ") _
         & " Per.COD_DEPARTAMENTO = '" & cboBuscaDept.ItemData(cboBuscaDept.ListIndex) & "'"
  vWhere = True
End If

If cboBuscaSec.Text <> "TODOS" And Len(cboBuscaSec.Text) > 0 Then
  strSQL = strSQL & IIf(vWhere, " AND ", " WHERE ") _
         & " Per.COD_SECCION = '" & cboBuscaSec.ItemData(cboBuscaSec.ListIndex) & "'"
  vWhere = True
End If




Call sbCargaGrid(vGridNomina, 7, strSQL)
vGridNomina.MaxRows = vGridNomina.MaxRows - 1

vPaso = False


Exit Sub

vError:
  vGridNomina.MaxRows = 0
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
    Call sbNomina
End If
End Sub

Private Sub sbBoletaResponsable(pIdentificacion As String, Optional pTipo As String = "B")
'Imprime la Boleta de Traslado

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de Activos Fijos"
 
 .Connect = glogon.ConectRPT

 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxSubTitulo = 'ACTIVOS VIGENTES'"
 
 If pTipo = "B" Then
    .ReportFileName = SIFGlobal.fxPathReportes("Activos_BoletaActivosAsignados.rpt")
 Else
    .ReportFileName = SIFGlobal.fxPathReportes("Activos_Contrato_Responsabilidad.rpt")
 End If
 
 .SelectionFormula = "{ACTIVOS_PERSONAS.IDENTIFICACION} = '" & pIdentificacion _
                   & "' AND {ACTIVOS_PRINCIPAL.ESTADO} <> 'R'"
  
 .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Sub sbBoletaTraslado(vBoleta As String)
'Imprime la Boleta de Traslado

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de Activos Fijos"
 
 .Connect = glogon.ConectRPT

 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxSubTitulo = 'TRASLADO DE ACTIVOS Y CAMBIO DE RESPONSABLES'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Activos_BoletaTraslado.rpt")
 .SelectionFormula = "{ACTIVOS_TRASLADOS.COD_TRASLADO} = '" & vBoleta & "'"
  
 .PrintReport

End With

Me.MousePointer = vbDefault

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Integer

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Línea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  strSQL = "delete Activos_Personas where identificacion = '" & vGrid.Text & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Persona : " & vGrid.Text)
    
  vGrid.DeleteRows vGrid.ActiveRow, 1
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGridNomina_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub


With vGridNomina

If Col = 2 Then
    fraCambio.Visible = True
    
    .Row = Row
    .Col = 4
    txtCedula.Text = .Text
    .Col = 5
    txtNombre.Text = .Text
   
End If

If Col = 3 Then
    .Row = Row
    .Col = 4
    Call sbBoletaResponsable(.Text)
End If


End With

End Sub
