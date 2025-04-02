VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Poliza_Proc_Recepcion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pólizas: Recepción y Facturación"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   15765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4815
      Left            =   0
      TabIndex        =   43
      Top             =   3480
      Width           =   15735
      _Version        =   1441793
      _ExtentX        =   27755
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
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1815
      Left            =   12120
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
      _Version        =   1441793
      _ExtentX        =   6165
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Datos del Asiento"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboUnidad 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3375
         _Version        =   1441793
         _ExtentX        =   5953
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
      Begin XtremeSuiteControls.ComboBox cboCentroCosto 
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   3375
         _Version        =   1441793
         _ExtentX        =   5953
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
         BackStyle       =   0  'Transparent
         Caption         =   "Centro Costo:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   240
      Width           =   5295
      _Version        =   1441793
      _ExtentX        =   9340
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   13800
      TabIndex        =   6
      Top             =   600
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Recepcion.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   375
      Left            =   14280
      TabIndex        =   7
      Top             =   600
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Recepcion.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   375
      Left            =   14760
      TabIndex        =   8
      Top             =   600
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Recepcion.frx":0E19
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   210
      Left            =   240
      TabIndex        =   32
      Top             =   3195
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
      Left            =   0
      TabIndex        =   33
      Top             =   8400
      Width           =   15735
      _Version        =   1441793
      _ExtentX        =   27755
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtCantidad 
         Height          =   315
         Left            =   4680
         TabIndex        =   34
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
         TabIndex        =   35
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
         TabIndex        =   36
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
         TabIndex        =   37
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
         Height          =   615
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   1085
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
         Picture         =   "frmCR_Poliza_Proc_Recepcion.frx":1532
      End
      Begin XtremeSuiteControls.PushButton btnFactura 
         Height          =   615
         Left            =   12480
         TabIndex        =   45
         Top             =   120
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Generar Factura en Cuentas por Pagar"
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
         Picture         =   "frmCR_Poliza_Proc_Recepcion.frx":1AD6
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
         TabIndex        =   42
         Top             =   120
         Width           =   1815
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
         TabIndex        =   41
         Top             =   480
         Width           =   1815
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
         TabIndex        =   40
         Top             =   120
         Width           =   2535
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
         TabIndex        =   39
         Top             =   480
         Width           =   2535
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   11895
      _Version        =   1441793
      _ExtentX        =   20981
      _ExtentY        =   3413
      _StockProps     =   79
      Caption         =   "Datos de la Factura"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboPoliza 
         Height          =   330
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.FlatEdit txtFactura 
         Height          =   435
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
         _ExtentY        =   767
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboProveedor 
         Height          =   330
         Left            =   1320
         TabIndex        =   14
         Top             =   1200
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1155
         Left            =   7680
         TabIndex        =   19
         Top             =   720
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   2037
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboProceso 
         Height          =   330
         Left            =   7680
         TabIndex        =   21
         Top             =   360
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   330
         Left            =   10440
         TabIndex        =   24
         Top             =   360
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
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   330
         Left            =   1320
         TabIndex        =   26
         Top             =   1560
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
      Begin XtremeSuiteControls.ComboBox cboFormaPago 
         Height          =   330
         Left            =   4560
         TabIndex        =   27
         Top             =   1560
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
      Begin XtremeSuiteControls.FlatEdit txtDivisaLocal 
         Height          =   330
         Left            =   4560
         TabIndex        =   30
         Top             =   840
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   582
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
         Text            =   "0.0"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTC 
         Height          =   330
         Left            =   3360
         TabIndex        =   31
         Top             =   1560
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   582
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
         Text            =   "1"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa Local:"
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
         Index           =   10
         Left            =   4560
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pago:"
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
         Index           =   9
         Left            =   4560
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa:"
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
         Index           =   8
         Left            =   -120
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vence:"
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
         Left            =   9480
         TabIndex        =   23
         Top             =   360
         Width           =   735
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
         Index           =   5
         Left            =   6600
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Notas:"
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
         Left            =   6240
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   -120
         TabIndex        =   13
         Top             =   1200
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
         Index           =   0
         Left            =   -120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
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
         Left            =   -120
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   3120
      Width           =   15735
      _Version        =   1441793
      _ExtentX        =   27755
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   8880
      TabIndex        =   5
      Top             =   240
      Width           =   975
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
      Caption         =   "Generación de Facturas de cuentas por pagar"
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
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmCR_Poliza_Proc_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean, vDivisa As String, vDivisaLocal As String
Dim vIVA_Porc As Currency, vIVA_Cta As String, vIVA_CtaDesc As String

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub cboUnidad_Click()
Dim strSQL As String

If vPaso Then Exit Sub
If cboUnidad.ListCount <= 0 Then Exit Sub


strSQL = "select RTRIM(COD_CENTRO_COSTO) as 'IdX', RTRIM(descripcion) as 'ItmX'" _
       & " From CNTX_CENTRO_COSTOS Where COD_CONTABILIDAD = " & GLOBALES.gEnlace & " And ACTIVO = 1" _
       & " and COD_CENTRO_COSTO in(select COD_CENTRO_COSTO  from CNTX_UNIDADES_CC" _
       & " where COD_CONTABILIDAD = " & GLOBALES.gEnlace & " and COD_UNIDAD = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "')"
vPaso = True
Call sbCbo_Llena_New(cboCentroCosto, strSQL, False, True)
vPaso = False

End Sub

Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Poliza", 1100
    .Add , , "No. Operacion", vbCenter
    .Add , , "Cédula", 2100, vbCenter
    .Add , , "Asegurado", 3500
    .Add , , "Monto Asegurado", 2500, vbRightJustify
    .Add , , "Monto Prima", 2500, vbRightJustify
End With


cboProceso.AddItem "202406"
cboProceso.Text = "202406"

strSQL = "select COD_POLIZA as 'IdX', DESCRIPCION as 'ItmX' From CRD_CATALOGO_POLIZAS"
Call sbCbo_Llena_New(cboPoliza, strSQL, False, True)

 'Carga la Divisa Local
 strSQL = "select rtrim(cod_divisa) as 'Divisa',rtrim(descripcion) as 'DivisaLocal' " _
        & " from CntX_Divisas where cod_contabilidad = " & GLOBALES.gEnlace _
        & " and Divisa_Local = 1"
 Call OpenRecordSet(rs, strSQL)
     vDivisa = rs!Divisa
     vDivisaLocal = rs!DivisaLocal
 rs.Close
 
 'Carga Divisas
 strSQL = "select rtrim(cod_divisa) as 'IdX',rtrim(descripcion) as 'ItmX'" _
        & " from CntX_Divisas where cod_contabilidad = " & GLOBALES.gEnlace _
        & " order by divisa_local desc,cod_divisa"
 Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)
 
 

 'Carga Unidades
 strSQL = "select rtrim(cod_unidad) as 'IdX',rtrim(descripcion) as 'ItmX'" _
        & " from CntX_unidades where cod_contabilidad = " & GLOBALES.gEnlace & " and activa = 1"
 Call sbCbo_Llena_New(cboUnidad, strSQL, False, True)

cboFormaPago.Clear
cboFormaPago.AddItem "Contado"
cboFormaPago.AddItem "Crédito"
cboFormaPago.Text = "Crédito"


End Sub

