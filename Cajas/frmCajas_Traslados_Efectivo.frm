VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCajas_Traslados_Efectivo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Traslado de Efectivo"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16860
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   16860
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   15975
      _Version        =   1441793
      _ExtentX        =   28178
      _ExtentY        =   10821
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
      Item(0).Caption =   "Pendientes"
      Item(0).Tooltip =   "Mis Traslados"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "btnResolucion(0)"
      Item(0).Control(2)=   "chkTodas"
      Item(0).Control(3)=   "gbAcciones"
      Item(0).Control(4)=   "btnResolucion(1)"
      Item(1).Caption =   "Historial"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "lswH"
      Item(1).Control(1)=   "cboHEstado"
      Item(1).Control(2)=   "cboHTipo"
      Item(1).Control(3)=   "dtpHInicio"
      Item(1).Control(4)=   "dtpHCorte"
      Item(1).Control(5)=   "chkHFechas"
      Item(1).Control(6)=   "btnHConsultar"
      Item(1).Control(7)=   "chkHTodas"
      Item(1).Control(8)=   "cboHFuente"
      Item(1).Control(9)=   "btnResolucion(2)"
      Begin XtremeSuiteControls.ListView lswH 
         Height          =   3615
         Left            =   -70000
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   13095
         _Version        =   1441793
         _ExtentX        =   23098
         _ExtentY        =   6376
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
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3855
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   13095
         _Version        =   1441793
         _ExtentX        =   23098
         _ExtentY        =   6800
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
      Begin XtremeSuiteControls.PushButton btnHConsultar 
         Height          =   375
         Left            =   -59800
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Consultar"
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
         Appearance      =   17
         Picture         =   "frmCajas_Traslados_Efectivo.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todas?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   375
         Index           =   1
         Left            =   11520
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Rechazar"
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
         Appearance      =   17
         Picture         =   "frmCajas_Traslados_Efectivo.frx":0700
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.CheckBox chkHTodas 
         Height          =   255
         Left            =   -69880
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todas?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   375
         Index           =   0
         Left            =   9960
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Aceptar"
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
         Appearance      =   17
         Picture         =   "frmCajas_Traslados_Efectivo.frx":0E16
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.ComboBox cboHEstado 
         Height          =   330
         Left            =   -66760
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.ComboBox cboHTipo 
         Height          =   330
         Left            =   -65200
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.DateTimePicker dtpHInicio 
         Height          =   330
         Left            =   -63040
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.DateTimePicker dtpHCorte 
         Height          =   330
         Left            =   -61720
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.CheckBox chkHFechas 
         Height          =   255
         Left            =   -60280
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   450
         _ExtentY        =   450
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   375
         Index           =   2
         Left            =   -58360
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         Appearance      =   17
         Picture         =   "frmCajas_Traslados_Efectivo.frx":153D
      End
      Begin XtremeSuiteControls.GroupBox gbAcciones 
         Height          =   1575
         Left            =   0
         TabIndex        =   24
         Top             =   4680
         Width           =   15735
         _Version        =   1441793
         _ExtentX        =   27755
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Realizar Traslado"
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
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
            Height          =   330
            Left            =   7320
            TabIndex        =   25
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   582
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboDMovimiento 
            Height          =   330
            Left            =   3960
            TabIndex        =   26
            Top             =   960
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtDCajaId 
            Height          =   330
            Left            =   480
            TabIndex        =   27
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.FlatEdit txtDCajaDesc 
            Height          =   330
            Left            =   1440
            TabIndex        =   28
            Top             =   600
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtImporte 
            Height          =   330
            Left            =   8520
            TabIndex        =   29
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboDivisa 
            Height          =   330
            Left            =   6120
            TabIndex        =   30
            Top             =   600
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
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
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   330
            Left            =   8520
            TabIndex        =   31
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   960
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
         Begin XtremeSuiteControls.PushButton btnRegistrar 
            Height          =   615
            Left            =   13920
            TabIndex        =   32
            Top             =   600
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Registrar"
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
            Appearance      =   17
            Picture         =   "frmCajas_Traslados_Efectivo.frx":1C3A
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   690
            Left            =   10320
            TabIndex        =   41
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Width           =   3375
            _Version        =   1441793
            _ExtentX        =   5953
            _ExtentY        =   1217
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   11
            Left            =   10320
            TabIndex        =   40
            Top             =   360
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notas"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   39
            Top             =   960
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Movimiento"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   38
            Top             =   360
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4895
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Descripción"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   37
            Top             =   360
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Caja Id"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   7
            Left            =   8520
            TabIndex        =   36
            Top             =   360
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Importe"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   8
            Left            =   6120
            TabIndex        =   35
            Top             =   360
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Divisa"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   9
            Left            =   7320
            TabIndex        =   34
            Top             =   360
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo Cambio"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   10
            Left            =   7320
            TabIndex        =   33
            Top             =   960
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto"
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
      Begin XtremeSuiteControls.ComboBox cboHFuente 
         Height          =   330
         Left            =   -68320
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
   End
   Begin XtremeSuiteControls.FlatEdit txtCajaId 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.FlatEdit txtCajaUsuario 
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   1200
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.FlatEdit txtCajaApertura 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1200
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.FlatEdit txtCajaDesc 
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   1200
      Width           =   6975
      _Version        =   1441793
      _ExtentX        =   12303
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   11760
      TabIndex        =   23
      Top             =   1200
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Appearance      =   17
      Picture         =   "frmCajas_Traslados_Efectivo.frx":235A
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   11760
      TabIndex        =   43
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   22
      Top             =   960
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Descripción"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   21
      Top             =   960
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Apertura Id"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   20
      Top             =   960
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Caja Id"
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
   Begin XtremeSuiteControls.Label lblTitle 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   9135
      _Version        =   1441793
      _ExtentX        =   16113
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Traslado de Efectivo entre Cajas"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmCajas_Traslados_Efectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean
Dim pCharRelleno As String, mTipoCambio As Currency

Private Sub sbCajaInicial()

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

pCharRelleno = fxCajasParametros("05")

Me.Caption = "Abonos a Créditos    ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtCajaId.Text = ModuloCajas.mCaja
txtCajaUsuario.Text = ModuloCajas.mUsuario
txtCajaApertura.Text = ModuloCajas.mApertura

strSQL = "select descripcion,ACTIVA, PERMITE_TRASLADOS_EF, ROL_BOVEDA  from cajas_definicion where cod_caja = '" & txtCajaId.Text & "'"
Call OpenRecordSet(rs, strSQL)
  txtCajaDesc.Text = rs!Descripcion

cboDMovimiento.Clear

If rs!PERMITE_TRASLADOS_EF = 0 And rs!ROL_BOVEDA = 0 Then
    cboDMovimiento.AddItem "Ninguno"
    cboDMovimiento.Text = "Ninguno"
End If

If rs!PERMITE_TRASLADOS_EF = 1 Then
   cboDMovimiento.AddItem "Traslado de Efectivo"
   cboDMovimiento.Text = "Traslado de Efectivo"
End If

If rs!ROL_BOVEDA = 1 Then
   cboDMovimiento.AddItem "Aprovisionamiento"
   cboDMovimiento.AddItem "Reintegro"
   cboDMovimiento.Text = "Aprovisionamiento"
End If

If rs!Activa = 0 Then
    cboDMovimiento.Clear
    cboDMovimiento.AddItem "Ninguno"
    cboDMovimiento.Text = "Ninguno"
End If

rs.Close

strSQL = "select RTRIM(COD_DIVISA) as 'Idx', rtrim(DESCRIPCION) AS 'ItmX'" _
      & " From vSys_Divisas" _
      & "  ORDER BY DIVISA_LOCAL DESC"
      
vPaso = True
    Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)
vPaso = False

Call cboDivisa_Click

Call sbConsulta


End Sub


Private Sub sbConsulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

lsw.ListItems.Clear

'create proc spCajas_TE_Consulta (@Caja varchar(10), @OrigenDestino char(1) = 'O' , @Movimiento varchar(10) = '', @Estado char(1) = ''
'        , @fInicio datetime = Null, @fCorte datetime = Null)
'TRASLADO_ID , Tipo, COD_DIVISA, COD_CAJA, COD_APERTURA, D_COD_CAJA, D_COD_APERTURA, REGISTRO_FECHA, REGISTRO_USUARIO
'    , ESTADO, TIPO_CAMBIO, MONTO, NOTAS
    
strSQL = "exec spCajas_TE_Consulta '" & txtCajaId.Text & "', 'D', '', 'P', Null, Null"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!TRASLADO_ID)
     itmX.SubItems(1) = rs!Cod_Caja
     itmX.SubItems(2) = rs!Registro_Usuario
     itmX.SubItems(3) = rs!Cod_Apertura
     itmX.SubItems(4) = rs!Registro_Fecha
     itmX.SubItems(5) = rs!cod_Divisa
     itmX.SubItems(6) = Format(rs!Importe, "Standard")
     itmX.SubItems(7) = rs!TIPO_CAMBIO
     itmX.SubItems(8) = Format(rs!Monto, "Standard")
     itmX.SubItems(9) = rs!NOTAS


 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbHistorico_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(1).Selected = True

lswH.ListItems.Clear

Dim pInicio As String, pCorte As String
    
If chkHFechas.Value = xtpChecked Then
   pInicio = "Null"
   pCorte = "Null"
Else
   pInicio = "'" & Format(dtpHInicio.Value, "yyyy-mm-dd") & " 00:00:00'"
   pCorte = "'" & Format(dtpHCorte.Value, "yyyy-mm-dd") & " 23:59:59'"

End If
    
strSQL = "exec spCajas_TE_Consulta '" & txtCajaId.Text & "', '" & Mid(cboHFuente.Text, 1, 1) _
        & "', '" & Mid(cboHTipo.Text, 1, 1) & "', '" & Mid(cboHEstado.Text, 1, 1) & "', " & pInicio & ", " & pCorte


Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lswH.ListItems.Add(, , rs!TRASLADO_ID)
     itmX.SubItems(1) = rs!Tipo_Descripcion
     itmX.SubItems(2) = rs!Cod_Caja
     itmX.SubItems(3) = rs!Registro_Usuario
     itmX.SubItems(4) = rs!Cod_Apertura
     itmX.SubItems(5) = rs!Registro_Fecha
     itmX.SubItems(6) = rs!cod_Divisa
     itmX.SubItems(7) = Format(rs!Importe, "Standard")
     itmX.SubItems(8) = rs!TIPO_CAMBIO
     itmX.SubItems(9) = Format(rs!Monto, "Standard")
     
     itmX.SubItems(10) = rs!Estado_Descripcion
     itmX.SubItems(11) = rs!D_COD_CAJA
     itmX.SubItems(12) = rs!ESTADO_USUARIO & ""
     itmX.SubItems(13) = rs!D_COD_APERTURA & ""
     itmX.SubItems(14) = rs!ESTADO_FECHA & ""
     itmX.SubItems(15) = rs!NOTAS

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Select Case tcMain.SelectedItem
 Case 0 'Pendientes
    Call Excel_Exportar_Lsw(lsw, ProgressBarX)
 Case 1 'Historico
    Call Excel_Exportar_Lsw(lswH, ProgressBarX)
End Select

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnHConsultar_Click()
If vPaso Then Exit Sub

 Call sbHistorico_Consulta
End Sub


Private Sub sbDivisaTipoCambio()
Dim vDivisa As String, i As Integer

If vPaso Then Exit Sub

'Cargar el Tipo de Cambio
mTipoCambio = 1

vDivisa = cboDivisa.ItemData(cboDivisa.ListIndex)

strSQL = "select dbo.fxCajas_TipoCambio(" & GLOBALES.gEnlace & ",'" & vDivisa & "',dbo.MyGetdate(),'C') as 'TipoCambio'"
Call OpenRecordSet(rs, strSQL)
  mTipoCambio = rs!TipoCambio
rs.Close

txtTipoCambio.Text = Format(mTipoCambio, "###,##0.0000")
txtMonto.Text = "0.00"
txtImporte.Text = "0.00"


End Sub

Private Function fxTraslado_Valida() As Boolean
Dim pResult As Boolean, pMensaje As String

pResult = True

pMensaje = ""

txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)

If cboDMovimiento.ListCount = 0 Then
    pMensaje = "- Esta Caja no está autorizada a realizar movimientos de traslado de efectivo" & vbCrLf
End If

If cboDMovimiento.Text = "Ninguno" Then
    pMensaje = "- Esta Caja no está autorizada a realizar movimientos de traslado de efectivo" & vbCrLf
End If

If txtCajaApertura.Text = "" Or txtCajaApertura.Text = "0" Then
    pMensaje = "- No existe una Caja abierta para iniciar proceso" & vbCrLf
End If

If txtDCajaId.Text = "" Then
    pMensaje = "- No se ha indicado una caja destino" & vbCrLf
End If

If CCur(txtImporte.Text) <= 0 Then
    pMensaje = "- El Importe a Registrar no es válido!" & vbCrLf
End If

If Len(pMensaje) > 0 Then
    pResult = False
    MsgBox pMensaje, vbExclamation
End If


fxTraslado_Valida = pResult

End Function

Private Sub btnRegistrar_Click()

On Error GoTo vError

If Not fxTraslado_Valida Then
    Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spCajas_TE_Registro '" & txtCajaId.Text & "', '" & txtCajaUsuario.Text & "', " & txtCajaApertura.Text _
       & ", '" & txtDCajaId.Text & "', '" & Mid(cboDMovimiento.Text, 1, 1) & "', '" & cboDivisa.ItemData(cboDivisa.ListIndex) _
       & "', " & CCur(txtTipoCambio.Text) & ", " & CCur(txtImporte.Text) & ", " & CCur(txtMonto.Text) _
       & ", '" & txtNotas.Text & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
       
Me.MousePointer = vbDefault

MsgBox "Tramite de " & cboDMovimiento.Text & " realizado satisfactoriamente!", vbInformation


txtDCajaId.Text = ""
txtDCajaDesc.Text = ""
txtNotas.Text = ""

txtImporte.Text = "0.00"
txtMonto.Text = "0.00"

Call sbConsulta


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnResolucion_Click(Index As Integer)

Dim i As Long, pResolucion As String

On Error GoTo vError

If txtCajaApertura.Text = "" Or txtCajaApertura.Text = "0" Then
    Exit Sub
End If

Me.MousePointer = vbHourglass


Select Case Index
    Case 0 'Aprobar
      pResolucion = "A"
    Case 1 'Rechazar
      pResolucion = "R"
    Case 2 'Cancelar
      pResolucion = "C"
End Select

' spCajas_TE_Resolucion (@TramiteId int , @Resolucion char(1), @Caja varchar(10), @CajaUsuario varchar(30), @CajaAperturaId int, @Usuario varchar(30))

If pResolucion = "C" Then
        With lswH.ListItems
        For i = 1 To .Count
          If .Item(i).Checked Then
                
                strSQL = "exec spCajas_TE_Resolucion " & .Item(i).Text & ", '" & pResolucion & "', '" & txtCajaId.Text _
                       & "', '" & txtCajaUsuario.Text & "', " & txtCajaApertura.Text & ", '" & glogon.Usuario & "'"
                Call ConectionExecute(strSQL)
         End If
        Next i

        End With
        
        Call sbHistorico_Consulta

Else

        With lsw.ListItems
        For i = 1 To .Count
          If .Item(i).Checked Then
                
                strSQL = "exec spCajas_TE_Resolucion " & .Item(i).Text & ", '" & pResolucion & "', '" & txtCajaId.Text _
                       & "', '" & txtCajaUsuario.Text & "', " & txtCajaApertura.Text & ", '" & glogon.Usuario & "'"
                Call ConectionExecute(strSQL)
         End If
        Next i
        
        End With

        Call sbConsulta

End If

       

Me.MousePointer = vbDefault

MsgBox "Resoluciones aplicadas satisfactoriamente!", vbInformation



Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboDivisa_Click()
If vPaso Then Exit Sub

Call sbDivisaTipoCambio

End Sub

Private Sub cboHEstado_Click()

If Mid(cboHEstado.Text, 1, 1) = "P" And Mid(cboHFuente.Text, 1, 1) = "O" Then
    lswH.Checkboxes = True
    chkHTodas.Visible = True
    btnResolucion(2).Visible = True
Else
    lswH.Checkboxes = False
    chkHTodas.Visible = False
    btnResolucion(2).Visible = False
End If

Call btnHConsultar_Click

End Sub

Private Sub cboHFuente_Click()
Call cboHEstado_Click
End Sub

Private Sub cboHTipo_Click()

Call btnHConsultar_Click

End Sub

Private Sub chkHFechas_Click()
If chkHFechas.Value = xtpChecked Then
    dtpHInicio.Enabled = False
    dtpHCorte.Enabled = False
Else
    dtpHInicio.Enabled = True
    dtpHCorte.Enabled = True
End If

End Sub


Private Sub chkHTodas_Click()
Dim i As Long

For i = 1 To lswH.ListItems.Count
    lswH.ListItems.Item(i).Checked = chkHTodas.Value
Next i

End Sub

Private Sub chkTodas_Click()
Dim i As Long

For i = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(i).Checked = chkTodas.Value
Next i

End Sub

Private Sub Form_Load()

vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

lblTitle.ForeColor = vbWhite

vPaso = True

cboHFuente.AddItem "Origen"
cboHFuente.AddItem "Destino"
cboHFuente.Text = "Origen"

cboHEstado.AddItem "Pendiente"
cboHEstado.AddItem "Aprobado"
cboHEstado.AddItem "Rechazado"
cboHEstado.AddItem "Cancelado"
cboHEstado.AddItem "TODOS"
cboHEstado.Text = "TODOS"

cboHTipo.AddItem "Aprovisionamientos"
cboHTipo.AddItem "Reintegros"
cboHTipo.AddItem "Traslados"
cboHTipo.AddItem "TODOS"
cboHTipo.Text = "TODOS"

dtpHInicio.Value = fxFechaServidor
dtpHCorte.Value = dtpHInicio.Value


With lsw.ColumnHeaders
    .Clear
    .Add , , "Tramite Id", 1400
    .Add , , "Caja Id", 1400, vbCenter
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Apertura Id", 1400, vbCenter
    .Add , , "Fecha", 2500
    .Add , , "Divisa", 1000, vbCenter
    .Add , , "Importe", 2500, vbRightJustify
    .Add , , "T.C.", 2500, vbRightJustify
    .Add , , "Monto", 2500, vbRightJustify
    .Add , , "Notas", 3500
End With

With lswH.ColumnHeaders
    .Clear
    .Add , , "Tramite Id", 1400
    .Add , , "Tipo", 2400
    .Add , , "Caja Id", 1400, vbCenter
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Apertura Id", 1400, vbCenter
    .Add , , "Fecha", 2500
    .Add , , "Divisa", 1000, vbCenter
    .Add , , "Importe", 2500, vbRightJustify
    .Add , , "T.C.", 2500, vbRightJustify
    .Add , , "Monto", 2500, vbRightJustify
    
    .Add , , "Estado", 1100, vbCenter
    .Add , , "Rec. Caja Id", 1400, vbCenter
    .Add , , "Rec. Usuario", 2500, vbCenter
    .Add , , "Rec. Apertura Id", 1400, vbCenter
    .Add , , "Rec. Fecha", 2500
    
    .Add , , "Notas", 3500
End With


vPaso = False

Call chkHFechas_Click

tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next


imgBanner.Width = Me.Width

tcMain.Width = Me.Width - 350
tcMain.Height = Me.Height - (tcMain.top + 450)

lsw.Left = 20
lswH.Left = 20

lsw.Width = tcMain.Width - 50
lsw.Height = tcMain.Height - (lsw.top + gbAcciones.Height + 200)
gbAcciones.top = lsw.top + lsw.Height + 100

gbAcciones.Width = lsw.Width


lswH.Width = tcMain.Width - 50
lswH.Height = tcMain.Height - (lswH.top + 150)
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call Form_Resize


End Sub

Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial

If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Unload Me
   Exit Sub
End If


End Sub

Private Sub txtDCajaId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "Select cod_caja,descripcion from cajas_definicion"
    gBusquedas.Columna = "cod_caja"
    gBusquedas.Orden = "cod_caja"
    gBusquedas.Filtro = " and cod_caja <> '" & txtCajaId.Text & "' and Activa = 1" _
                      & " and (PERMITE_TRASLADOS_EF = 1 or ROL_BOVEDA = 1)"
    frmBusquedas.Show vbModal
    txtDCajaId.Text = Trim(gBusquedas.Resultado)
    txtDCajaDesc.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtImporte_GotFocus()
On Error GoTo vError
    txtImporte.Text = CCur(txtImporte.Text)
vError:
End Sub

Private Sub txtImporte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub


Private Sub txtImporte_LostFocus()
On Error GoTo vError
    txtImporte.Text = Format(CCur(txtImporte.Text), "Standard")
    Call sbMonto_Calculo
vError:

End Sub

Private Sub sbMonto_Calculo()

If Not IsNumeric(txtImporte.Text) Then
    MsgBox "Debe digitar solamente números...", vbExclamation
    txtImporte.Text = 0
    txtMonto.Text = 0
Else
    txtMonto.Text = Format(CCur(txtImporte.Text) * fxSys_Tipo_Cambio_Apl(mTipoCambio), "Standard")
End If

End Sub
