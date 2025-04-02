VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Poliza_Proc_Prevista 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pólizas: Prevista de Recepción"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7695
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   12135
      _Version        =   1441793
      _ExtentX        =   21405
      _ExtentY        =   13573
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
      Item(0).Caption =   "Generación - Prevista"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "btnBuscar"
      Item(0).Control(1)=   "btnCargar"
      Item(0).Control(2)=   "btnInfo"
      Item(0).Control(3)=   "txtArchivo"
      Item(0).Control(4)=   "Label1(2)"
      Item(0).Control(5)=   "txtR_Factura"
      Item(0).Control(6)=   "Label2(7)"
      Item(0).Control(7)=   "Label2(0)"
      Item(0).Control(8)=   "cboPoliza"
      Item(0).Control(9)=   "cboProceso"
      Item(0).Control(10)=   "Label2(1)"
      Item(0).Control(11)=   "lsw"
      Item(0).Control(12)=   "btnGenerar"
      Item(0).Control(13)=   "gbResumen"
      Item(0).Control(14)=   "ShortcutCaption1"
      Item(0).Control(15)=   "chkTodos"
      Item(1).Caption =   "Consulta"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "Label2(2)"
      Item(1).Control(1)=   "Label2(3)"
      Item(1).Control(2)=   "btnC_Exportar"
      Item(1).Control(3)=   "cboC_Poliza"
      Item(1).Control(4)=   "cboC_Proceso"
      Item(1).Control(5)=   "btnC_Buscar"
      Item(1).Control(6)=   "lswC"
      Begin XtremeSuiteControls.ListView lswC 
         Height          =   6615
         Left            =   -69880
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   11668
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4815
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   8493
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   210
         Left            =   360
         TabIndex        =   20
         Top             =   1520
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox gbResumen 
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   6720
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   1720
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtCantidad 
            Height          =   315
            Left            =   4680
            TabIndex        =   21
            Top             =   120
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTotal 
            Height          =   315
            Left            =   4680
            TabIndex        =   22
            Top             =   480
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSelCantidad 
            Height          =   315
            Left            =   9000
            TabIndex        =   23
            Top             =   120
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSelMonto 
            Height          =   315
            Left            =   9000
            TabIndex        =   24
            Top             =   480
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnEliminar 
            Height          =   495
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Eliminar Seleccionados"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Poliza_Proc_Prevista.frx":0000
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Seleccionado:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   6240
            TabIndex        =   18
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Registros Seleccionados:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   6240
            TabIndex        =   17
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MontoTotal:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   16
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Total:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   15
            Top             =   120
            Width           =   1815
         End
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Left            =   8760
         TabIndex        =   3
         Top             =   960
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Proc_Prevista.frx":05A4
      End
      Begin XtremeSuiteControls.PushButton btnCargar 
         Height          =   375
         Left            =   9240
         TabIndex        =   4
         Top             =   960
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Proc_Prevista.frx":0CA4
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   375
         Left            =   9720
         TabIndex        =   5
         Top             =   960
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Proc_Prevista.frx":13BD
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12086
         _ExtentY        =   656
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
      Begin XtremeSuiteControls.FlatEdit txtR_Factura 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboPoliza 
         Height          =   330
         Left            =   6120
         TabIndex        =   11
         Top             =   480
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.PushButton btnGenerar 
         Height          =   375
         Left            =   10320
         TabIndex        =   13
         Top             =   960
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Generar Prevista"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Proc_Prevista.frx":1AD6
      End
      Begin XtremeSuiteControls.ComboBox cboC_Poliza 
         Height          =   330
         Left            =   -68560
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.ComboBox cboC_Proceso 
         Height          =   330
         Left            =   -64360
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.PushButton btnC_Buscar 
         Height          =   375
         Left            =   -62440
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Proc_Prevista.frx":1C40
      End
      Begin XtremeSuiteControls.PushButton btnC_Exportar 
         Height          =   375
         Left            =   -61960
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_Poliza_Proc_Prevista.frx":2340
      End
      Begin XtremeSuiteControls.ComboBox cboProceso 
         Height          =   330
         Left            =   10320
         TabIndex        =   33
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proceso:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Proceso:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -65680
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Póliza:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -69760
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Resultados"
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
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Póliza:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Polizas de Vivienda y Prendario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prevista - Auxiliar de Pólizas"
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
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "frmCR_Poliza_Proc_Prevista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Poliza", 1100
    .Add , , "No. Operacion", vbCenter
    .Add , , "Cédula", 2100, vbCenter
    .Add , , "Asegurado", 3500
    .Add , , "Monto Asegurado", 2500, vbRightJustify
    .Add , , "Monto Prima", 2500, vbRightJustify
End With

With lswC.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Cédula", 1000, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Póliza", 1100, vbCenter
    .Add , , "Proceso", 1100, vbCenter
    .Add , , "Id Envío", 2000, vbCenter
    .Add , , "Monto Envío", 2100, vbRightJustify
    .Add , , "Id Recibo", 2000, vbCenter
    .Add , , "Monto Recibo", 2100, vbRightJustify
    .Add , , "Diferencia", 2100, vbRightJustify
    .Add , , "Factura", 2000
    .Add , , "No.Operación", 1500, vbCenter
    .Add , , "Tipo Match", 2000, vbCenter
    
End With


cboProceso.AddItem "202406"
cboProceso.Text = "202406"

strSQL = "select COD_POLIZA as 'IdX', DESCRIPCION as 'ItmX' From CRD_CATALOGO_POLIZAS"
Call sbCbo_Llena_New(cboPoliza, strSQL, False, True)

Call sbCbo_Copia(cboPoliza, cboC_Poliza)
Call sbCbo_Copia(cboProceso, cboC_Proceso)


End Sub

