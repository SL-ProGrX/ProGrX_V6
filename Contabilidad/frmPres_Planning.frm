VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmPres_Planning 
   Caption         =   "Presupuesto: Vista Mensual"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkCtaMov 
      Height          =   210
      Left            =   9720
      TabIndex        =   26
      Top             =   360
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
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
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   11280
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbAjustes 
      Height          =   5415
      Left            =   2760
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   8055
      _Version        =   1441793
      _ExtentX        =   14203
      _ExtentY        =   9546
      _StockProps     =   79
      Caption         =   "Ajuste de Partida   "
      ForeColor       =   16711680
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
      Begin XtremeSuiteControls.PushButton btnAjuste 
         Height          =   615
         Left            =   5280
         TabIndex        =   15
         Top             =   4680
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Ajustar"
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
         Picture         =   "frmPres_Planning.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnAjuste_Cancela 
         Height          =   615
         Left            =   6480
         TabIndex        =   16
         Top             =   4680
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   1085
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
         Picture         =   "frmPres_Planning.frx":07DE
      End
      Begin XtremeSuiteControls.ComboBox cboAjuste 
         Height          =   312
         Left            =   1800
         TabIndex        =   32
         Top             =   1560
         Width           =   6012
         _Version        =   1441793
         _ExtentX        =   10610
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtAJ_Ajuste 
         Height          =   312
         Left            =   1800
         TabIndex        =   39
         Top             =   2400
         Width           =   1932
         _Version        =   1441793
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAJ_ValorActual 
         Height          =   312
         Left            =   1800
         TabIndex        =   40
         Top             =   2040
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtAJ_ValorNuevo 
         Height          =   312
         Left            =   1800
         TabIndex        =   41
         Top             =   2760
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtAJ_Disponible 
         Height          =   312
         Left            =   5880
         TabIndex        =   42
         Top             =   2040
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtAJ_Acumulado 
         Height          =   312
         Left            =   5880
         TabIndex        =   43
         Top             =   2400
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtAJ_PresupuestoTotal 
         Height          =   312
         Left            =   5880
         TabIndex        =   44
         Top             =   2760
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtAJ_Notas 
         Height          =   912
         Left            =   1800
         TabIndex        =   45
         Top             =   3480
         Width           =   6012
         _Version        =   1441793
         _ExtentX        =   10604
         _ExtentY        =   1609
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAJ_Cuenta 
         Height          =   792
         Left            =   1800
         TabIndex        =   46
         Top             =   600
         Width           =   6012
         _Version        =   1441793
         _ExtentX        =   10604
         _ExtentY        =   1397
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAJ_PresupuestoTotalNuevo 
         Height          =   312
         Left            =   5880
         TabIndex        =   47
         Top             =   3120
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   550
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
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Presupuesto Nuevo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   15
         Left            =   4080
         TabIndex        =   48
         Top             =   3120
         Width           =   1692
      End
      Begin VB.Label lblAJ_CentroCosto 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "[...]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   4920
         TabIndex        =   36
         Top             =   240
         Width           =   2892
      End
      Begin VB.Label lblAJ_Unidad 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "[...]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1800
         TabIndex        =   35
         Top             =   240
         Width           =   3012
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
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
         Height          =   252
         Index           =   7
         Left            =   4080
         TabIndex        =   34
         Top             =   2040
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Ajuste"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   33
         Top             =   1560
         Width           =   2412
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   14
         Left            =   360
         TabIndex        =   24
         Top             =   3480
         Width           =   2412
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Presupuesto Total"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   12
         Left            =   4080
         TabIndex        =   14
         Top             =   2760
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Acumulado Actual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   11
         Left            =   4080
         TabIndex        =   13
         Top             =   2400
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensual Nuevo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   10
         Left            =   360
         TabIndex        =   12
         Top             =   2760
         Width           =   2412
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   9
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensual Actual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   360
         TabIndex        =   10
         Top             =   2040
         Width           =   2412
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   2412
      End
   End
   Begin XtremeSuiteControls.GroupBox gbDetalle 
      Height          =   2772
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   12612
      _Version        =   1441793
      _ExtentX        =   22246
      _ExtentY        =   4890
      _StockProps     =   79
      Caption         =   "Detalle de la Cuenta"
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
      Begin XtremeSuiteControls.ListView lswHistorico 
         Height          =   2172
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9123
         _ExtentY        =   3831
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
      Begin XtremeSuiteControls.ListView lswPresupuesto 
         Height          =   2172
         Left            =   5640
         TabIndex        =   23
         Top             =   600
         Width           =   7092
         _Version        =   1441793
         _ExtentX        =   12509
         _ExtentY        =   3831
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
      Begin XtremeSuiteControls.PushButton btnCopy 
         Height          =   252
         Index           =   0
         Left            =   12120
         TabIndex        =   37
         Top             =   240
         Width           =   252
         _Version        =   1441793
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPres_Planning.frx":0FAB
      End
      Begin XtremeSuiteControls.RadioButton rbDetalle 
         Height          =   252
         Index           =   0
         Left            =   5520
         TabIndex        =   29
         Top             =   240
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Presupuesto vrs Real"
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
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbDetalle 
         Height          =   252
         Index           =   1
         Left            =   7680
         TabIndex        =   30
         Top             =   240
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ajustes del Periodo"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbDetalle 
         Height          =   255
         Index           =   2
         Left            =   9960
         TabIndex        =   31
         Top             =   240
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todos los Ajustes"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnCopy 
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   38
         Top             =   240
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPres_Planning.frx":1115
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro Histórico Real:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2412
      End
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   852
      Left            =   11760
      TabIndex        =   4
      Top             =   360
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Appearance      =   17
      Picture         =   "frmPres_Planning.frx":127F
      TextImageRelation=   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4452
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   12732
      _Version        =   524288
      _ExtentX        =   22458
      _ExtentY        =   7853
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
      RowHeaderDisplay=   0
      SpreadDesigner  =   "frmPres_Planning.frx":1C9D
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   5640
      TabIndex        =   17
      Top             =   720
      Width           =   3972
      _Version        =   1441793
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.ComboBox cboCentroCosto 
      Height          =   312
      Left            =   5640
      TabIndex        =   18
      Top             =   1080
      Width           =   3972
      _Version        =   1441793
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.ComboBox cboModelo 
      Height          =   312
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6588
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
   Begin XtremeSuiteControls.ComboBox cboContabilidad 
      Height          =   312
      Left            =   4080
      TabIndex        =   20
      Top             =   360
      Width           =   5532
      _Version        =   1441793
      _ExtentX        =   9763
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
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   312
      Left            =   240
      TabIndex        =   25
      Top             =   960
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6588
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
   Begin XtremeSuiteControls.CheckBox chkRealEnTramite 
      Height          =   210
      Left            =   9720
      TabIndex        =   27
      ToolTipText     =   "Se utiliza para traer el dato contable más cernano aún cuando no se encuentre cerrado el periodo o aplicado el movimiento."
      Top             =   840
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   852
      Left            =   12720
      TabIndex        =   28
      Top             =   360
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Appearance      =   17
      Picture         =   "frmPres_Planning.frx":2929
      TextImageRelation=   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   492
      Index           =   1
      Left            =   9960
      TabIndex        =   50
      ToolTipText     =   "No muestra cuentas de resumen, solo las que reciben movimientos"
      Top             =   720
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Calcular Real en Trámite?"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   492
      Index           =   0
      Left            =   9960
      TabIndex        =   49
      ToolTipText     =   "No muestra cuentas de resumen, solo las que reciben movimientos"
      Top             =   240
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Sin Cuentas de Resumen?"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
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
      Height          =   252
      Index           =   13
      Left            =   4080
      TabIndex        =   21
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo "
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
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
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
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Centro de Costo"
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
      Height          =   252
      Index           =   6
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Negocio"
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
      Height          =   252
      Index           =   5
      Left            =   4080
      TabIndex        =   0
      Top             =   720
      Width           =   1692
   End
   Begin XtremeShortcutBar.ShortcutCaption scBanner 
      Height          =   1455
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   14175
      _Version        =   1441793
      _ExtentX        =   25003
      _ExtentY        =   2566
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmPres_Planning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim vModelo As String, vCierreID As Long, vModeloAbierto As Boolean
Dim mCuenta As String, mUnidad As String, mCentroCosto As String


Private Sub btnAjuste_Cancela_Click()
gbAjustes.Visible = False
End Sub

Private Sub btnAjuste_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Dim pContabilidad As Long, pModelo As String, pTipoAjuste As String
Dim pUnidad As String, pCentroCosto As String, pCuenta As String, pCtaMov As String
Dim iMes As Integer, iAnio As Integer

On Error GoTo vError

pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)

pTipoAjuste = cboAjuste.ItemData(cboAjuste.ListIndex)

iAnio = Year(CDate(cboPeriodo.Text))
iMes = Month(CDate(cboPeriodo.Text))

pUnidad = lblAJ_Unidad.Tag
pCentroCosto = lblAJ_CentroCosto.Tag
pCuenta = txtAJ_Cuenta.Tag


'Validar el Tipo de Ajustes +/- y que la cuenta reciba movimientos

strSQL = "SELECT * FROM pres_tipos_ajustes where cod_ajuste = '" & pTipoAjuste & "'"
Call OpenRecordSet(rs, strSQL)

If CCur(txtAJ_Ajuste.Text) > 0 And rs!AJUSTE_LIBRE_POSITIVO = 0 Then
    MsgBox "El tipo de Ajuste: " & UCase(cboAjuste.Text) & " ,no concuerda con el valor del cambio!", vbExclamation
    Exit Sub
End If

If CCur(txtAJ_Ajuste.Text) < 0 And rs!AJUSTE_LIBRE_NEGATIVO = 0 Then
    MsgBox "El tipo de Ajuste: " & UCase(cboAjuste.Text) & " ,no concuerda con el valor del cambio!", vbExclamation
    Exit Sub
End If

rs.Close


'Validar Movimientos sobre Consolidados
If pUnidad = "CONSOLIDADO" Or pCentroCosto = "CONSOLIDADO" Then
    MsgBox "No se permiten ajustes sobre Consolidados!", vbExclamation
    Exit Sub
End If


Me.MousePointer = vbHourglass


'ALTER PROCEDURE spPres_PresupuestoAjustesGuarda(
'        @Contabilidad       int,
'        @Modelo             varchar(10),
'        @Anio               int,
'        @Mes                smallint,
'        @Cuenta             varchar(60),
'        @Mnt_MensualNuevo   decimal(18,2),
'        @Mnt_Ajuste         decimal(18,2),
'        @Unidad             varchar(10),
'        @CentroCosto        varchar(10),
'        @Notas              varchar(1000),
'        @Usuario            varchar(30)
        
'Aplicar Ajustes
strSQL = "exec spPres_PresupuestoAjustesGuarda " & pContabilidad & ",'" _
       & pModelo & "'," & iAnio & "," & iMes & ",'" & pCuenta & "'," _
       & CCur(txtAJ_ValorNuevo.Text) & "," & CCur(txtAJ_Ajuste.Text) _
       & ",'" & pUnidad & "','" & pCentroCosto & "','" & txtAJ_Notas.Text _
       & "','" & glogon.Usuario & "','" & cboAjuste.ItemData(cboAjuste.ListIndex) & "'"
       
Call ConectionExecute(strSQL)
       
Me.MousePointer = vbDefault

MsgBox "Ajustes aplicados satisfactoriamente!", vbInformation

Call btnBuscar_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub btnBuscar_Click()
Dim i As Long, x As Integer, strSQL As String, rs As New ADODB.Recordset
Dim pContabilidad As Long, pModelo As String
Dim pUnidad As String, pCentroCosto As String, pCuenta As String, pCtaMov As String
Dim iMes As Integer, iAnio As Integer
Dim vVista As String


On Error GoTo vError

Me.MousePointer = vbHourglass

gbAjustes.Visible = False

If chkCtaMov.Value = xtpChecked Then
  pCtaMov = "1"
Else
  pCtaMov = "Null"
End If

pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)
pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

If cboCentroCosto.ListCount = 0 Then
  pCentroCosto = ""
Else
  pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
End If


iAnio = Year(CDate(cboPeriodo.Text))
iMes = Month(CDate(cboPeriodo.Text))


If cboUnidad.Text = "CONSOLIDADO" And cboCentroCosto.Text = "CONSOLIDADO" Then
  vVista = "G"
Else
  vVista = "U"
  If cboCentroCosto.Text <> "CONSOLIDADO" Then
    vVista = "C"
  End If
End If 'Consolidado

        
strSQL = "exec spPres_VistaPresupuesto " & pContabilidad & ",'" _
       & pModelo & "','" & pUnidad & "','" & pCentroCosto & "'," _
       & iAnio & "," & iMes & ",'" & vVista & "'," & pCtaMov

With vGrid
    .MaxRows = 0
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 .MaxRows = .MaxRows + 1
 .Row = .MaxRows
      
      
 .FontBold = False
 
 .Col = 3
 .Text = rs!Cuenta
 .CellTag = rs!Acepta_movimientos
 .Col = 4
 .Text = rs!Descripcion
 
 
 .Col = 5
 .Text = rs!Cod_Unidad & ""
 .Col = 6
 .Text = rs!cod_Centro_Costo & ""
 
 .Col = 7
 .Text = Format(rs!REAL_MES, "Standard")
 .Col = 8
 .Text = Format(rs!MENSUAL, "Standard")
 .Col = 9
 .Text = Format(rs!DIFERENCIA_MES, "Standard")
 .Col = 10
 .Text = Format(rs!REAL_ACUMULADO, "Standard")
 .Col = 11
 .Text = Format(rs!acumulado, "Standard")
 .Col = 12
 .Text = Format(rs!DIFERENCIA_ACUMULADA, "Standard")
 .Col = 13
 .Text = Format(rs!Pres_Total, "Standard")
 .Col = 14
 .Text = Format(rs!DIFERENCIA_TOTAL, "Standard")
 
 
 .Col = 15
 .Text = Format(rs!EJECUTADO_MES * 100, "Standard")
 .Col = 16
 .Text = Format(rs!EJECUTADO_ACUMULADO * 100, "Standard")
 .Col = 17
 .Text = Format(rs!EJECUTADO_TOTAL * 100, "Standard")
  
  
  If rs!Acepta_movimientos = 0 Then
         For x = 3 To .MaxCols
             .Col = x
             .FontBold = True
         Next x
  End If
  
 rs.MoveNext
Loop
rs.Close

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCopy_Click(Index As Integer)
Dim i As Long, j As Long
Dim pTrayIcon As XtremeSuiteControls.TrayIcon


On Error GoTo vError
 

Set pTrayIcon = frmContenedor.TrayIcon
 
On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Index
    Case 0
        Call Excel_Exportar_Lsw(lswPresupuesto)
    
    Case 1 'Historico Real
        Call Excel_Exportar_Lsw(lswHistorico)
            
End Select

 
pTrayIcon.ShowBalloonTip 25, "ProGrX: Notificación" _
            , "Exportación a Excel concluida" _
            , xtpToolTipIconInfo


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
Dim vArchivo As String, Ex As Boolean

    vHeaders.Columnas = 15

    vHeaders.Headers(1) = "[+/-]"
    vHeaders.Headers(2) = "[...]"
    vHeaders.Headers(3) = "Cuenta"
    vHeaders.Headers(4) = "Descripción"
    vHeaders.Headers(5) = "Unidad"
    vHeaders.Headers(6) = "Centro"
    
    vHeaders.Headers(7) = "Real Mensual"
    vHeaders.Headers(8) = "Presupuesto Mensual"
    vHeaders.Headers(9) = "Diferencia Mensual"
    vHeaders.Headers(10) = "Real Acumulado"
    vHeaders.Headers(11) = "Presupuesto Acumulado"
    vHeaders.Headers(12) = "Diferencia Acumulada"
    
    vHeaders.Headers(13) = "(% Ejecución Mensual)"
    vHeaders.Headers(14) = "(% Ejecución Acumulada)"
    vHeaders.Headers(15) = "(% Ejecución Total)"
    
 
vArchivo = "Presupuesto_" & cboModelo.ItemData(cboModelo.ListIndex) & "_" & Format(cboPeriodo.Text, "yyyy-mm-dd")
Call sbSIFGridExportar(vGrid, vHeaders, vArchivo, "Excel")

End Sub

Private Sub cboCentroCosto_Click()
If vPaso Then Exit Sub

vGrid.MaxRows = 0
gbAjustes.Visible = False

End Sub

Private Sub cboContabilidad_Click()
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

vPaso = True


strSQL = "select P.cod_modelo as 'IdX' , P.DESCRIPCION as 'ItmX', Cc.Inicio_Anio" _
       & " From PRES_MODELOS P INNER JOIN PRES_MODELOS_USUARIOS Pmu on P.cod_Contabilidad = Pmu.cod_contabilidad" _
       & "  and P.cod_Modelo = Pmu.cod_Modelo and Pmu.Usuario = '" & glogon.Usuario & "'" _
       & " INNER JOIN CNTX_CIERRES Cc on P.cod_Contabilidad = Cc.cod_Contabilidad and P.ID_CIERRE = Cc.ID_CIERRE " _
       & " Where P.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & " group by P.cod_Modelo, P.Descripcion, Cc.Inicio_Anio" _
       & " order by Cc.INICIO_ANIO desc, P.Cod_Modelo"
Call sbCbo_Llena_New(cboModelo, strSQL, False, True)

vPaso = False

Call cboModelo_Click

Exit Sub

vError:
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboModelo_Click()
If vPaso Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If cboModelo.ListCount = 0 Then Exit Sub

vModeloAbierto = False

strSQL = "select Cc.INICIO_ANIO,Cc.INICIO_MES, Cc.CORTE_ANIO, Cc.CORTE_MES, Pm.Estado" _
       & " from CNTX_CIERRES Cc inner join PRES_MODELOS Pm on Cc.COD_CONTABILIDAD = Pm.COD_CONTABILIDAD and Cc.ID_CIERRE = Pm.ID_CIERRE" _
       & " where Pm.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & " and Pm.COD_MODELO = '" & cboModelo.ItemData(cboModelo.ListIndex) & "'" _
       & " order by Cc.INICIO_ANIO desc"
Call OpenRecordSet(rs, strSQL)

If rs!Estado = "P" Then
    vModeloAbierto = True
End If

strSQL = "select dbo.fxSys_FechaAnioMesToDatetime(anio,mes) as 'ItmX'" _
       & " From dbo.fxPres_Periodo(" & rs!Inicio_Anio & "," & rs!Inicio_Mes & "," & rs!Corte_Anio & "," & rs!Corte_Mes & "," & cboContabilidad.ItemData(cboContabilidad.ListIndex) & ")"
rs.Close

'Call sbCbo_Llena_New(cboPeriodo, strSQL, False, False)

cboPeriodo.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboPeriodo.AddItem rs!itmX & ""

 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboPeriodo.Text = rs!itmX & ""
End If
rs.Close

vPaso = True

strSQL = "exec spPres_Modelo_Unidades " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & ",'" & cboModelo.ItemData(cboModelo.ListIndex) & "','" & glogon.Usuario & "'"

Call sbCbo_Llena_New(cboUnidad, strSQL, False, True)

If cboUnidad.ListCount > 0 Then
    cboUnidad.AddItem "CONSOLIDADO"
    cboUnidad.ItemData(cboUnidad.ListCount - 1) = "CONSOLIDADO"
    
    cboUnidad.Text = "CONSOLIDADO"
End If

strSQL = "exec spPres_Modelo_Ajustes_Permitidos " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & ",'" & cboModelo.ItemData(cboModelo.ListIndex) & "','" & glogon.Usuario & "'"

Call sbCbo_Llena_New(cboAjuste, strSQL, False, True)

vPaso = False
Call cboUnidad_Click

vGrid.MaxRows = 0
gbAjustes.Visible = False
Exit Sub

vError:
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cboPeriodo_Click()

vGrid.MaxRows = 0
gbAjustes.Visible = False

End Sub

Private Sub cboUnidad_Click()
If vPaso Then Exit Sub


Dim strSQL As String, pUnidad As String

On Error GoTo vError

vPaso = True


If cboUnidad.Text = "CONSOLIDADO" Then
   pUnidad = "CONS"
Else
   pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)
End If

strSQL = "EXEc spPres_Modelo_Unidades_CC " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & ",'" & cboModelo.ItemData(cboModelo.ListIndex) & "','" & pUnidad & "'"
Call sbCbo_Llena_New(cboCentroCosto, strSQL, True, True)


If cboCentroCosto.ListCount > 0 Then
    cboCentroCosto.AddItem "CONSOLIDADO"
    cboCentroCosto.ItemData(cboCentroCosto.ListCount - 1) = "CONSOLIDADO"
    
    cboCentroCosto.Text = "CONSOLIDADO"
End If


vPaso = False


vGrid.MaxRows = 0
gbAjustes.Visible = False

Exit Sub

vError:
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 12

lswHistorico.ColumnHeaders.Clear
lswHistorico.ColumnHeaders.Add , , "Periodo", 1240
lswHistorico.ColumnHeaders.Add , , "Mensual", 1640, vbRightJustify
lswHistorico.ColumnHeaders.Add , , "Acumulado", 1840, vbRightJustify


lswPresupuesto.ColumnHeaders.Clear
lswPresupuesto.ColumnHeaders.Add , , "Periodo", 1240
lswPresupuesto.ColumnHeaders.Add , , "Presupuesto", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "Real", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "Diferencia", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "( % )", 900, vbCenter
lswPresupuesto.ColumnHeaders.Add , , "( + ) Ajuste", 2240, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "( - ) Ajuste", 2240, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "Original", 2440, vbRightJustify

vPaso = True

cboPeriodo.AddItem "Enero", 1
cboPeriodo.AddItem "Febrero", 2
cboPeriodo.AddItem "Marzo", 3
cboPeriodo.AddItem "Abril", 4
cboPeriodo.AddItem "Mayo", 5
cboPeriodo.AddItem "Junio", 6
cboPeriodo.AddItem "Julio", 7
cboPeriodo.AddItem "Agosto", 8
cboPeriodo.AddItem "Setiembre", 9
cboPeriodo.AddItem "Octubre", 10
cboPeriodo.AddItem "Noviembre", 11
cboPeriodo.AddItem "Diciembre", 12
cboPeriodo.Text = "Enero"

vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error GoTo vError

vGrid.Width = Me.Width - 350
vGrid.Height = Me.Height - (vGrid.Top + gbDetalle.Height + 800)

gbDetalle.Top = vGrid.Height + vGrid.Top + 150
gbDetalle.Width = vGrid.Width

scBanner.Width = Me.Width

lswPresupuesto.Width = gbDetalle.Width - (lswPresupuesto.Left + 150)

vError:
End Sub


Private Sub lswHistorico_DblClick()

If vPaso Or lswHistorico.ListItems.Count = 0 Then Exit Sub

Dim pCuenta As String

pCuenta = gbDetalle.Tag

gCntX_Presupuesto.Cuenta = pCuenta
gCntX_Presupuesto.Periodo = Format(lswHistorico.SelectedItem.Text, "yyyy/mm/dd")
gCntX_Presupuesto.Modelo = cboModelo.ItemData(cboModelo.ListIndex)
gCntX_Presupuesto.Contabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
gCntX_Presupuesto.Unidad = cboUnidad.ItemData(cboUnidad.ListIndex)
gCntX_Presupuesto.Centro = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)

Dim frm As Form

Call sbFormActivo("frmPres_Analitico", frm)
If frm Is Nothing Then
Else
    UnLoad frm
End If

Call sbFormsCall("frmPres_Analitico", , , , False, Me)
End Sub

Private Sub lswPresupuesto_DblClick()
If vPaso Or lswPresupuesto.ListItems.Count = 0 Then Exit Sub

Dim pCuenta As String


If rbDetalle.Item(0).Value Then
    pCuenta = gbDetalle.Tag
    
    
    
    gCntX_Presupuesto.Cuenta = pCuenta
    gCntX_Presupuesto.Periodo = Format(lswPresupuesto.SelectedItem.Text, "yyyy/mm/dd")
    gCntX_Presupuesto.Modelo = cboModelo.ItemData(cboModelo.ListIndex)
    gCntX_Presupuesto.Contabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
    gCntX_Presupuesto.Unidad = lswPresupuesto.SelectedItem.SubItems(10)
    gCntX_Presupuesto.Centro = lswPresupuesto.SelectedItem.SubItems(11)
    
    Dim frm As Form
    
    Call sbFormActivo("frmPres_Analitico", frm)
    If frm Is Nothing Then
    Else
        UnLoad frm
    End If
    Call sbFormsCall("frmPres_Analitico", , , , False, Me)
End If


End Sub

Private Sub rbDetalle_Click(Index As Integer)
Dim i As Long, x As Integer, strSQL As String, rs As New ADODB.Recordset
Dim pContabilidad As Long, pModelo As String
Dim pUnidad As String, pCentroCosto As String, pPeriodo As String, pCuenta As String
Dim iMes As Integer, iAnio As Integer
Dim vVista As String, itmX As ListViewItem


On Error GoTo vError

Me.MousePointer = vbHourglass

pCuenta = gbDetalle.Tag
pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)

'pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)
'
'
'If cboCentroCosto.ListCount = 0 Then
'  pCentroCosto = ""
'Else
'  pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
'End If

pUnidad = mUnidad
pCentroCosto = mCentroCosto


iAnio = Year(CDate(cboPeriodo.Text))
iMes = Month(CDate(cboPeriodo.Text))

        
If pUnidad = "CONSOLIDADO" And pCentroCosto = "CONSOLIDADO" Then
  vVista = "G"
Else
  vVista = "U"
  If pCentroCosto <> "CONSOLIDADO" Then
    vVista = "C"
  End If
End If 'Consolidado
        

rbDetalle.Item(Index).Value = True

With lswPresupuesto
    .ListItems.Clear



Dim curAcumulado As Currency, curReal_Acumulado As Currency, curDiferencia_Acumulada As Currency, curEjecutado_Acumulado As Currency
Dim curAjustePositivo As Currency, curAjusteNegativo As Currency, curPreMensualInicial As Currency
Dim curDiferenciaTotal As Currency, curEjecutadoTotal As Currency


curAcumulado = 0
curReal_Acumulado = 0
curDiferencia_Acumulada = 0
curEjecutado_Acumulado = 0
curAjustePositivo = 0
curAjusteNegativo = 0
curDiferenciaTotal = 0
curEjecutadoTotal = 0
curPreMensualInicial = 0


Select Case Index
   Case 0 'Presupuesto del Periodo
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Periodo", 1240
        .ColumnHeaders.Add , , "Presupuesto", 2440, vbRightJustify
        .ColumnHeaders.Add , , "Real", 2440, vbRightJustify
        .ColumnHeaders.Add , , "Diferencia", 2440, vbRightJustify
        .ColumnHeaders.Add , , "( % )", 900, vbCenter
        .ColumnHeaders.Add , , "( + ) Ajuste", 2240, vbRightJustify
        .ColumnHeaders.Add , , "( - ) Ajuste", 2240, vbRightJustify
        .ColumnHeaders.Add , , "Original", 2440, vbRightJustify
        
        .ColumnHeaders.Add , , "Unidad", 900, vbCenter
        .ColumnHeaders.Add , , "Centro", 900, vbCenter
        
        strSQL = "exec spPres_VistaPresupuesto_Cuenta " & pContabilidad & ",'" _
               & pModelo & "','" & pUnidad & "','" & pCentroCosto & "','" & pCuenta & "','" & vVista & "'"
            
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Set itmX = .ListItems.Add(, , rs!Periodo)
              itmX.SubItems(1) = Format(rs!acumulado, "Standard")
              itmX.SubItems(2) = Format(rs!REAL_ACUMULADO, "Standard")
              itmX.SubItems(3) = Format(rs!DIFERENCIA_ACUMULADA, "Standard")
              itmX.SubItems(4) = Format(rs!EJECUTADO_ACUMULADO * 100, "Standard")
              itmX.SubItems(5) = Format(rs!AJUSTE_POSITIVO, "Standard")
              itmX.SubItems(6) = Format(rs!AJUSTE_NEGATIVO, "Standard")
              itmX.SubItems(7) = Format(rs!PRE_MENSUAL_INICIAL, "Standard")
         
              itmX.SubItems(8) = rs!Cod_Unidad & ""
              itmX.SubItems(9) = rs!cod_Centro_Costo & ""
         
         
            curAcumulado = curAcumulado + rs!acumulado
            curReal_Acumulado = curReal_Acumulado + rs!REAL_ACUMULADO
            curDiferencia_Acumulada = curDiferencia_Acumulada + rs!DIFERENCIA_ACUMULADA
            curEjecutado_Acumulado = 0 'EJECUTADO_ACUMULADO * 100
            curAjustePositivo = curAjustePositivo + rs!AJUSTE_POSITIVO
            curAjusteNegativo = curAjusteNegativo + rs!AJUSTE_NEGATIVO
            curDiferenciaTotal = curDiferenciaTotal + rs!DIFERENCIA_TOTAL
            curEjecutadoTotal = 0 'rs!EJECUTADO_TOTAL * 100
            curPreMensualInicial = curPreMensualInicial + rs!PRE_MENSUAL_INICIAL
         rs.MoveNext
        Loop
        rs.Close

      Set itmX = .ListItems.Add(, , "TOTALES:")
          itmX.SubItems(1) = Format(curAcumulado, "Standard")
          itmX.SubItems(2) = Format(curReal_Acumulado, "Standard")
          itmX.SubItems(3) = Format(curDiferencia_Acumulada, "Standard")
          itmX.SubItems(4) = Format(curEjecutado_Acumulado, "Standard")
          itmX.SubItems(5) = Format(curAjustePositivo, "Standard")
          itmX.SubItems(6) = Format(curAjusteNegativo, "Standard")
          itmX.SubItems(7) = Format(curPreMensualInicial, "Standard")
         
         
  
   Case 1, 2 'Ajustes
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Periodo", 1240
        .ColumnHeaders.Add , , "I.Acumulado", 2440, vbRightJustify
        .ColumnHeaders.Add , , "I.Mensual", 2440, vbRightJustify
        .ColumnHeaders.Add , , "( + ) Ajuste", 2240, vbRightJustify
        .ColumnHeaders.Add , , "F.Acumulado", 2440, vbRightJustify
        .ColumnHeaders.Add , , "F.Mensual", 2440, vbRightJustify
        .ColumnHeaders.Add , , "Fecha", 2240
        .ColumnHeaders.Add , , "Usuario", 2240
        
        .ColumnHeaders.Add , , "Unidad", 900, vbCenter
        .ColumnHeaders.Add , , "Centro", 900, vbCenter
        
        
'spPres_PresupuestoAjustesConsulta](
'    @Contabilidad       int,
'    @Modelo             varchar(10),
'    @Unidad             varchar(10) = Null,
'    @CentroCosto        varchar(10) = Null,
'    @Cuenta             varchar(30) = Null,
'    @Periodo            datetime    = Null
       If Index = 1 Then
            strSQL = "exec spPres_PresupuestoAjustesConsulta " & pContabilidad & ",'" _
                   & pModelo & "','" & pUnidad & "','" & pCentroCosto & "','" & pCuenta & "','" & Format(cboPeriodo.Text, "yyyy/mm/dd") & "'"
       Else
            strSQL = "exec spPres_PresupuestoAjustesConsulta " & pContabilidad & ",'" _
                   & pModelo & "','" & pUnidad & "','" & pCentroCosto & "','" & pCuenta & "',Null"
       End If
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Set itmX = .ListItems.Add(, , rs!Periodo)
'              itmX.SubItems(1) = Format(rs!ACUMULADO, "Standard")
'              itmX.SubItems(2) = Format(rs!REAL_ACUMULADO, "Standard")
'              itmX.SubItems(3) = Format(rs!DIFERENCIA_ACUMULADA, "Standard")
'              itmX.SubItems(4) = Format(rs!EJECUTADO_ACUMULADO * 100, "Standard")
'              itmX.SubItems(5) = Format(rs!AJUSTE_POSITIVO, "Standard")
'              itmX.SubItems(6) = Format(rs!AJUSTE_NEGATIVO, "Standard")
'              itmX.SubItems(7) = Format(rs!PRE_MENSUAL_INICIAL, "Standard")
         rs.MoveNext
        Loop
        rs.Close
            
End Select

        
        

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0
TimerX.Enabled = False

 vPaso = True '
    strSQL = "select cod_contabilidad as 'IdX', Nombre as 'ItmX' from CNTX_Contabilidades" _
           & " order by cod_contabilidad"
    Call sbCbo_Llena_New(cboContabilidad, strSQL, False, True)
 vPaso = False
 
 Call cboContabilidad_Click
 
End Sub

Private Sub sbHistorico(pCuenta As String)
Dim i As Long, x As Integer, strSQL As String, rs As New ADODB.Recordset
Dim pContabilidad As Long, pModelo As String
Dim pUnidad As String, pCentroCosto As String
Dim iMes As Integer, iAnio As Integer
Dim vVista As String, itmX As ListViewItem


On Error GoTo vError

Me.MousePointer = vbHourglass


pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)


pUnidad = mUnidad
pCentroCosto = mCentroCosto


iAnio = Year(CDate(cboPeriodo.Text))
iMes = Month(CDate(cboPeriodo.Text))

If pUnidad = "CONSOLIDADO" And pCentroCosto = "CONSOLIDADO" Then
  vVista = "G"
Else
  vVista = "U"
  If pCentroCosto <> "CONSOLIDADO" Then
    vVista = "C"
  End If
End If 'Consolidado
        
        
'Presupuesto del Periodo
lswPresupuesto.ColumnHeaders.Clear
lswPresupuesto.ColumnHeaders.Add , , "Periodo", 1240
lswPresupuesto.ColumnHeaders.Add , , "Presupuesto", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "Real", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "Diferencia", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "( % )", 900, vbCenter
lswPresupuesto.ColumnHeaders.Add , , "( + ) Ajuste", 2240, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "( - ) Ajuste", 2240, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "Original", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "Dif.Total", 2440, vbRightJustify
lswPresupuesto.ColumnHeaders.Add , , "( % Total )", 1000, vbCenter
lswPresupuesto.ColumnHeaders.Add , , "Unidad", 1000, vbCenter
lswPresupuesto.ColumnHeaders.Add , , "Centro", 1000, vbCenter


strSQL = "exec spPres_VistaPresupuesto_Cuenta " & pContabilidad & ",'" _
       & pModelo & "','" & pUnidad & "','" & pCentroCosto & "','" & pCuenta & "','" & vVista & "'"
       
       
Dim curAcumulado As Currency, curReal_Acumulado As Currency, curDiferencia_Acumulada As Currency, curEjecutado_Acumulado As Currency
Dim curAjustePositivo As Currency, curAjusteNegativo As Currency, curPreMensualInicial As Currency
Dim curDiferenciaTotal As Currency, curEjecutadoTotal As Currency


curAcumulado = 0
curReal_Acumulado = 0
curDiferencia_Acumulada = 0
curEjecutado_Acumulado = 0
curAjustePositivo = 0
curAjusteNegativo = 0
curDiferenciaTotal = 0
curEjecutadoTotal = 0
curPreMensualInicial = 0

With lswPresupuesto
    .ListItems.Clear
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , Format(rs!Periodo, "yyyy-mm-dd"))
          itmX.SubItems(1) = Format(rs!acumulado, "Standard")
          itmX.SubItems(2) = Format(rs!REAL_ACUMULADO, "Standard")
          itmX.SubItems(3) = Format(rs!DIFERENCIA_ACUMULADA, "Standard")
          itmX.SubItems(4) = Format(rs!EJECUTADO_ACUMULADO * 100, "Standard")
          itmX.SubItems(5) = Format(rs!AJUSTE_POSITIVO, "Standard")
          itmX.SubItems(6) = Format(rs!AJUSTE_NEGATIVO, "Standard")
          itmX.SubItems(7) = Format(rs!PRE_MENSUAL_INICIAL, "Standard")
          itmX.SubItems(8) = Format(rs!DIFERENCIA_TOTAL, "Standard")
          itmX.SubItems(9) = Format(rs!EJECUTADO_TOTAL * 100, "Standard")
     
          itmX.SubItems(10) = rs!Cod_Unidad & ""
          itmX.SubItems(11) = rs!cod_Centro_Costo & ""
     
     
            curAcumulado = rs!acumulado
            curReal_Acumulado = rs!REAL_ACUMULADO
            curDiferencia_Acumulada = rs!DIFERENCIA_ACUMULADA
            curEjecutado_Acumulado = 0 'EJECUTADO_ACUMULADO * 100
            curAjustePositivo = curAjustePositivo + rs!AJUSTE_POSITIVO
            curAjusteNegativo = curAjusteNegativo + rs!AJUSTE_NEGATIVO
            curDiferenciaTotal = curDiferenciaTotal + rs!DIFERENCIA_TOTAL
            curEjecutadoTotal = 0 'rs!EJECUTADO_TOTAL * 100
            curPreMensualInicial = curPreMensualInicial + rs!PRE_MENSUAL_INICIAL
     rs.MoveNext
    Loop
    rs.Close

      Set itmX = .ListItems.Add(, , "TOTALES:")
          itmX.SubItems(1) = Format(curAcumulado, "Standard")
          itmX.SubItems(2) = Format(curReal_Acumulado, "Standard")
          itmX.SubItems(3) = Format(curDiferencia_Acumulada, "Standard")
          itmX.SubItems(4) = Format(curEjecutado_Acumulado, "Standard")
          itmX.SubItems(5) = Format(curAjustePositivo, "Standard")
          itmX.SubItems(6) = Format(curAjusteNegativo, "Standard")
          itmX.SubItems(7) = Format(curPreMensualInicial, "Standard")
          itmX.SubItems(8) = Format(curDiferenciaTotal, "Standard")
          itmX.SubItems(9) = Format(curEjecutadoTotal, "Standard")

End With



'Historico Real
strSQL = "exec spPres_Cuenta_Real_Historico " & pContabilidad & ",'" _
       & pModelo & "'," & Month(CDate(cboPeriodo.Text)) & ",'" & pUnidad & "','" & pCentroCosto _
       & "','" & pCuenta & "','" & vVista & "'"

With lswHistorico
    .ListItems.Clear
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , Format(rs!Periodo, "yyyy-mm-dd"))
          itmX.SubItems(1) = Format(rs!NETO_MES, "Standard")
          itmX.SubItems(2) = Format(rs!saldo_final, "Standard")
     rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbAjustes_Load(pCuenta As String)
Dim i As Long, x As Integer, strSQL As String, rs As New ADODB.Recordset
Dim pContabilidad As Long, pModelo As String
Dim pUnidad As String, pCentroCosto As String
Dim vVista As String


On Error GoTo vError


If Not vModeloAbierto Then
    gbAjustes.Visible = True
Else
    gbAjustes.Visible = False
    Exit Sub
End If

Me.MousePointer = vbHourglass


pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)

pUnidad = lblAJ_Unidad.Tag
pCentroCosto = lblAJ_CentroCosto.Tag


If pUnidad = "CONSOLIDADO" Then
        vVista = "G"
ElseIf pCentroCosto = "CONSOLIDADO" Then
    vVista = "U"
ElseIf pCentroCosto = "TODOS" Then
    vVista = "C"
Else
    vVista = "D"
End If


strSQL = "exec spPres_VistaPresupuesto_Cuenta " & pContabilidad & ",'" _
       & pModelo & "','" & pUnidad & "','" & pCentroCosto & "','" & pCuenta _
       & "','" & vVista & "','" & Format(cboPeriodo.Text, "yyyy/MM/dd hh:mm:ss") & "'"
       
Call OpenRecordSet(rs, strSQL)
      
txtAJ_Acumulado.Text = Format(rs!acumulado, "Standard")
txtAJ_PresupuestoTotal.Text = Format(rs!Pres_Total, "Standard")
txtAJ_Disponible.Text = Format(rs!DIFERENCIA_ACUMULADA, "Standard")

txtAJ_ValorActual.Text = Format(rs!MENSUAL, "Standard")
txtAJ_ValorActual.ToolTipText = "Presupuesto Inicial: " & Format(rs!PRE_MENSUAL_INICIAL, "Standard")

txtAJ_Ajuste.Text = Format(0, "Standard")

txtAJ_ValorNuevo.Text = Format(CCur(txtAJ_ValorActual.Text) + CCur(txtAJ_Ajuste.Text), "Standard")
txtAJ_PresupuestoTotalNuevo.Text = Format(CCur(txtAJ_PresupuestoTotal.Text) + CCur(txtAJ_Ajuste.Text), "Standard")

rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtAJ_Ajuste_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtAJ_ValorNuevo.SetFocus
End If
End Sub

Private Sub txtAJ_Ajuste_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

txtAJ_ValorNuevo.Text = Format(CCur(txtAJ_ValorActual.Text) + CCur(txtAJ_Ajuste.Text), "Standard")
txtAJ_PresupuestoTotalNuevo.Text = Format(CCur(txtAJ_PresupuestoTotal.Text) + CCur(txtAJ_Ajuste.Text), "Standard")
Exit Sub

vError:
End Sub

Private Sub txtAJ_Ajuste_LostFocus()
On Error GoTo vError

txtAJ_Ajuste.Text = Format(CCur(txtAJ_Ajuste.Text), "Standard")
Exit Sub

vError:
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim pCuenta As String, pPreTotal As Currency
Dim pUnidad As String, pCentroCosto As String


vGrid.Row = Row

vGrid.Col = 5
pUnidad = vGrid.Text


vGrid.Col = 6
pCentroCosto = vGrid.Text


If pUnidad = "-C-" Then
    pUnidad = "CONSOLIDADO"
End If

If pCentroCosto = "-C-" Then
    pCentroCosto = "CONSOLIDADO"
End If

vGrid.Col = 3
pCuenta = fxgCntCuentaFormato(False, vGrid.Text, 0)

gbDetalle.Caption = "Detalle de la Cuenta" & Space(20) & "[" & vGrid.Text & "] [U: " & pUnidad & " ] [Cc: " & pCentroCosto & " ]" _


vGrid.Col = 11
pPreTotal = vGrid.Text

vGrid.Col = 4
gbDetalle.Caption = gbDetalle.Caption & Space(10) & vGrid.Text & Space(10) & "[Presupuesto Total:" _
        & Format(pPreTotal, "Standard") & "]"



gbDetalle.Tag = pCuenta

mCuenta = pCuenta
mUnidad = pUnidad
mCentroCosto = pCentroCosto

Select Case Col
    Case 1 'Ajustes
        vGrid.Col = 3
        'Acepta Movimientos
        If vGrid.CellTag = 1 And Not vModeloAbierto Then
            txtAJ_Cuenta.Tag = pCuenta
            vGrid.Col = 3
            txtAJ_Cuenta.Text = vGrid.Text & vbCrLf
            
            vGrid.Col = 4
            txtAJ_Cuenta.Text = txtAJ_Cuenta.Text & vGrid.Text & vbCrLf & cboPeriodo.Text
             
            lblAJ_Unidad.Caption = "Unidad: " & pUnidad
            lblAJ_CentroCosto.Caption = "Centro Costo: " & pCentroCosto
                         
            lblAJ_Unidad.Tag = pUnidad
            lblAJ_CentroCosto.Tag = pCentroCosto
                         
            Call sbAjustes_Load(pCuenta)
                         
        Else
          MsgBox "La cuenta no acepta movimientos o el modelo presupuestario no está autorizado para ajustes!", vbInformation
        End If
        
    Case 2 'Historicos
       Call sbHistorico(pCuenta)
End Select


End Sub
