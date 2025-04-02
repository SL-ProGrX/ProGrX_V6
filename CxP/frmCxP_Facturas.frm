VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCxPFacturas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Facturas (No Comerciales)"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   12030
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   315
      Left            =   8640
      TabIndex        =   54
      Top             =   0
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAux 
      Height          =   330
      Index           =   0
      Left            =   4320
      TabIndex        =   51
      Top             =   0
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Anular"
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
      Picture         =   "frmCxP_Facturas.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnAux 
      Height          =   330
      Index           =   1
      Left            =   5520
      TabIndex        =   52
      Top             =   0
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Plantilla"
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
      Picture         =   "frmCxP_Facturas.frx":05A4
   End
   Begin XtremeSuiteControls.PushButton btnAux 
      Height          =   330
      Index           =   2
      Left            =   6720
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Info"
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
      Picture         =   "frmCxP_Facturas.frx":0CAC
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   315
      Left            =   9240
      TabIndex        =   37
      Top             =   480
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3408
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
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   3
   End
   Begin XtremeSuiteControls.GroupBox gbFactura 
      Height          =   1095
      Left            =   11280
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   11775
      _Version        =   1572864
      _ExtentX        =   20770
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Cambio de Número de Factura:"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnCambio 
         Height          =   315
         Left            =   6600
         TabIndex        =   36
         Top             =   480
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Reemplazar No. Factura"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtFacturaNew 
         Height          =   312
         Left            =   3720
         TabIndex        =   35
         Top             =   480
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo, No. Factura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   13
         Left            =   1920
         TabIndex        =   34
         Top             =   480
         Width           =   1692
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtFactura 
      Height          =   315
      Left            =   1320
      TabIndex        =   29
      Top             =   480
      Width           =   3375
      _Version        =   1572864
      _ExtentX        =   5953
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbDatos 
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   11775
      _Version        =   1572864
      _ExtentX        =   20770
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Datos de la Factura:"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkCargosFlotantesAplica 
         Height          =   264
         Left            =   7440
         TabIndex        =   26
         Top             =   960
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "Aplica Cargos?"
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
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnImpuesto 
         Height          =   312
         Left            =   4344
         TabIndex        =   21
         Top             =   960
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Actualiza"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   7440
         TabIndex        =   22
         Top             =   600
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
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   312
         Left            =   2520
         TabIndex        =   23
         Top             =   600
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   312
         Left            =   9120
         TabIndex        =   28
         Top             =   600
         Width           =   1332
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
      Begin XtremeSuiteControls.FlatEdit txtImpuesto 
         Height          =   312
         Left            =   2520
         TabIndex        =   40
         Top             =   960
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
         Height          =   312
         Left            =   4320
         TabIndex        =   42
         Top             =   600
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisaLocal 
         Height          =   312
         Left            =   5520
         TabIndex        =   41
         Top             =   600
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
         Height          =   315
         Left            =   480
         TabIndex        =   44
         Top             =   600
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Divisa Local"
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
         Height          =   252
         Index           =   10
         Left            =   5520
         TabIndex        =   39
         Top             =   360
         Width           =   1932
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Vence"
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
         Height          =   252
         Index           =   5
         Left            =   9120
         TabIndex        =   20
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Divisa"
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
         Index           =   3
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Forma..Pago"
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
         Left            =   7440
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total a Pagar"
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
         Index           =   11
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipo de Cambio"
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
         Height          =   252
         Index           =   9
         Left            =   4320
         TabIndex        =   16
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "I.V.:"
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
         Height          =   252
         Index           =   15
         Left            =   1920
         TabIndex        =   15
         Top             =   960
         Width           =   492
      End
   End
   Begin XtremeSuiteControls.GroupBox gbAsiento 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   11775
      _Version        =   1572864
      _ExtentX        =   20770
      _ExtentY        =   6376
      _StockProps     =   79
      Caption         =   "Asiento:"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   3360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxP_Facturas.frx":13C5
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxP_Facturas.frx":14FC
               Key             =   "Copiar"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxP_Facturas.frx":1603
               Key             =   "Asiento"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxP_Facturas.frx":1715
               Key             =   "Info"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxP_Facturas.frx":1839
               Key             =   "verde"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxP_Facturas.frx":1957
               Key             =   "amarillo"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxP_Facturas.frx":1A7D
               Key             =   "rojo"
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   11535
         _Version        =   524288
         _ExtentX        =   20346
         _ExtentY        =   3625
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
         MaxCols         =   491
         ScrollBars      =   2
         SpreadDesigner  =   "frmCxP_Facturas.frx":1BA7
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnPlantilla 
         Height          =   315
         Left            =   10680
         TabIndex        =   13
         ToolTipText     =   "Carga Planilla"
         Top             =   480
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   16777215
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
      Begin XtremeSuiteControls.ComboBox cboUnidad 
         Height          =   312
         Left            =   720
         TabIndex        =   24
         Top             =   480
         Width           =   4332
         _Version        =   1572864
         _ExtentX        =   7646
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
      Begin XtremeSuiteControls.ComboBox cboCentroCosto 
         Height          =   312
         Left            =   5040
         TabIndex        =   25
         Top             =   480
         Width           =   3972
         _Version        =   1572864
         _ExtentX        =   7011
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
      Begin XtremeSuiteControls.FlatEdit txtPlantilla 
         Height          =   312
         Left            =   9360
         TabIndex        =   43
         Top             =   480
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
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
      Begin XtremeSuiteControls.FlatEdit txtDebito 
         Height          =   315
         Left            =   8040
         TabIndex        =   45
         Top             =   3120
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCredito 
         Height          =   315
         Left            =   9840
         TabIndex        =   46
         Top             =   3120
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiferencia 
         Height          =   315
         Left            =   1440
         TabIndex        =   47
         Top             =   3120
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia"
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
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   1092
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Totales"
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
         Index           =   0
         Left            =   6840
         TabIndex        =   11
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Unidad"
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
         Index           =   7
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   8
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   1932
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plantilla"
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
         Height          =   252
         Index           =   12
         Left            =   9360
         TabIndex        =   8
         Top             =   240
         Width           =   732
      End
   End
   Begin ComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7470
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   9596
            MinWidth        =   9596
            Text            =   "Registrado por:"
            TextSave        =   "Registrado por:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   6244
            MinWidth        =   6244
            Text            =   "Saldo: "
            TextSave        =   "Saldo: "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.CheckBox chkPlantilla 
      Height          =   270
      Left            =   6240
      TabIndex        =   27
      Top             =   480
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   466
      _StockProps     =   79
      Caption         =   "Plantilla?"
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   675
      Left            =   1320
      TabIndex        =   31
      Top             =   1320
      Width           =   9870
      _Version        =   1572864
      _ExtentX        =   17410
      _ExtentY        =   1191
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
   Begin XtremeSuiteControls.FlatEdit txtProvDivisa 
      Height          =   315
      Left            =   9720
      TabIndex        =   32
      Top             =   960
      Width           =   1470
      _Version        =   1572864
      _ExtentX        =   2593
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
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   315
      Left            =   9240
      TabIndex        =   38
      Top             =   480
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3408
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProvCod 
      Height          =   312
      Left            =   1320
      TabIndex        =   30
      Top             =   960
      Width           =   1224
      _Version        =   1572864
      _ExtentX        =   2159
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
   Begin XtremeSuiteControls.FlatEdit txtProvDesc 
      Height          =   315
      Left            =   2520
      TabIndex        =   48
      Top             =   960
      Width           =   7215
      _Version        =   1572864
      _ExtentX        =   12726
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   11400
      TabIndex        =   49
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   0
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   582
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCxP_Facturas.frx":22C6
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
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
   Begin VB.Image imgCambio 
      Height          =   255
      Left            =   5760
      Picture         =   "frmCxP_Facturas.frx":234F
      Stretch         =   -1  'True
      ToolTipText     =   "Cambio de No. Factura"
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgEstado 
      Height          =   255
      Left            =   5400
      Picture         =   "frmCxP_Facturas.frx":2CA4
      Stretch         =   -1  'True
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Factura"
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
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
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   8280
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmCxPFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vPaso As Boolean, vDivisa As String, vDivisaLocal As String
Dim vIVA_Porc As Currency, vIVA_Cta As String, vIVA_CtaDesc As String



Private Sub btnAdjuntos_Click()
 gGA.Modulo = "CXP"
 gGA.Llave_01 = txtProvCod.Text
 gGA.Llave_02 = txtFactura.Text
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub btnAux_Click(Index As Integer)
Dim i As Byte

Select Case Index
 Case 0 'Anular
    If Mid(txtEstado.Text, 1, 1) = "A" Then
       MsgBox "Esta factura ya se encuentra Anulada!", vbInformation
       Exit Sub
    End If
    
    If Mid(txtEstado.Text, 1, 1) = "P" Then
        i = MsgBox("Esta seguro que desea Anular la factura: " & txtFactura.Text & " ...?", vbYesNo)
        If i = vbYes Then Call sbAnular
    End If
 
 Case 1 'Plantillas
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "cod_factura"
    gBusquedas.Orden = "cod_factura"
    gBusquedas.Consulta = "select F.cod_factura,P.descripcion as Proveedor,F.total,F.notas" _
                  & " from cxp_facturas F inner join cxp_proveedores P on F.cod_proveedor = P.cod_proveedor"
    gBusquedas.Filtro = " and plantilla = 1"
    frmBusquedas.Show vbModal
    txtFactura = gBusquedas.Resultado
    If txtFactura <> "" Then
      Call sbConsulta(gBusquedas.Resultado)
      txtFactura = ""
      txtFactura.SetFocus
      chkPlantilla.Value = vbUnchecked
    End If
    
   Case 2 'Información
    MsgBox "No se localiza información de Proceso!", vbInformation
End Select
End Sub

Private Sub btnCambio_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spCxP_Factura_Cambio_No " & txtProvCod.Text & ",'" & txtFactura.Text _
    & "','" & txtFacturaNew.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If Not glogon.error Then
    Call Bitacora("Aplica", "Cambio Factura:" & txtFactura.Text & " -> " & txtFacturaNew.Text & ", Prov.Id:" & txtProvCod.Text)
    
    MsgBox "Cambio de No. Factura realizado satisfactoriamente!", vbInformation
    Call sbConsulta(txtFacturaNew.Text)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnImpuesto_Click()
Dim strSQL As String

On Error GoTo vError

If vEdita Then
   strSQL = "update cxp_facturas set IMPUESTO_VENTAS = " & CCur(txtImpuesto.Text) _
          & " where cod_proveedor = " & txtProvCod.Text & " and cod_factura = '" & txtFactura.Text & "'"
   Call ConectionExecute(strSQL)
   
   
   Call Bitacora("Modifica", "CxP-Factura: " & txtFactura & "...Prov:" & txtProvCod.Text & ", IV:" & txtImpuesto.Text)

    MsgBox "Monto de Impuesto Actualizado!", vbInformation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbInformation

End Sub


Private Sub btnPlantilla_Click()
 
If Len(txtPlantilla.Text) <> 0 Then
 Call sbPlantillaAsiento
End If
 
End Sub

Private Sub cboCentroCosto_Click()

If vPaso Then Exit Sub
If cboCentroCosto.ListCount <= 0 Then Exit Sub

Call sbCambiaInfo

End Sub

Private Sub cboCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub


Private Sub sbCalculoDivisaLocal()
On Error GoTo vError

If IsNumeric(txtTotalPagar.Text) Then
   txtDivisaLocal.Text = Format(CCur(txtTotalPagar.Text) * CCur(txtTipoCambio.Text), "Standard")
End If

vError:

End Sub

Private Sub cboDivisa_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub
If cboDivisa.ListCount <= 0 Then Exit Sub

On Error GoTo vError

strSQL = "select dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "','" & Format(dtpFecha.Value, "yyyy-MM-dd") & "','V') as Tipo_Cambio"
Call OpenRecordSet(rs, strSQL)
  txtTipoCambio.Text = Format(rs!TIPO_CAMBIO, "###,###.00####")
rs.Close

If CCur(txtTipoCambio.Text) <> 1 Then
   txtTipoCambio.Locked = False
Else
   txtTipoCambio.Locked = True
End If


Call sbCalculoDivisaLocal
Call sbCambiaInfo


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboDivisa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboUnidad.SetFocus
End Sub

Private Sub cboTipo_Click()
If cboTipo.ListCount = 0 Then Exit Sub

If cboTipo.Text = "Crédito" Then
   chkCargosFlotantesAplica.Value = vbChecked
   chkCargosFlotantesAplica.Visible = False
Else
   chkCargosFlotantesAplica.Value = vbChecked
   chkCargosFlotantesAplica.Visible = True
End If

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVence.SetFocus
End Sub

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

Call sbCambiaInfo

End Sub


Private Sub cboUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCentroCosto.SetFocus
End Sub

Private Sub dtpFecha_Change()
Call cboDivisa_Click
End Sub

Private Sub dtpVence_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDivisa.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_factura from cxp_facturas"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_factura > '" & txtFactura.Text & "'"
       
       If txtProvCod.Text <> "" And IsNumeric(txtProvCod.Text) Then
          strSQL = strSQL & " and cod_proveedor = " & txtProvCod.Text & " order by cod_factura asc"
       Else
          strSQL = strSQL & " order by cod_factura asc"
       End If
    Else
       strSQL = strSQL & " where cod_factura < '" & txtFactura.Text & "'"
       If txtProvCod.Text <> "" And IsNumeric(txtProvCod.Text) Then
          strSQL = strSQL & " and cod_proveedor = " & txtProvCod.Text & " order by cod_factura desc"
       Else
          strSQL = strSQL & " order by cod_factura desc"
       End If
    
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_Factura)
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
vModulo = 30
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError
 
 vModulo = 30
 
 vGrid.AppearanceStyle = fxGridStyle
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vPaso = True
 
 vIVA_Porc = 0.13
 vIVA_Cta = ""
 vIVA_CtaDesc = ""
 
 strSQL = "select COD_PARAMETRO, VALOR , DESCRIPCION " _
        & " From SYS_IVA_PARAMETROS" _
        & " where COD_PARAMETRO in('02','03','08')"
 Call OpenRecordSet(rs, strSQL)
 Do While Not rs.EOF
  
  Select Case Trim(rs!Cod_Parametro)
    Case "03" 'Soportado no Identificado
       vIVA_Cta = RTrim(rs!Valor)
       vIVA_CtaDesc = fxgCntCuentaDesc(vIVA_Cta)
    Case "08" 'Porcentaje Default IVA
      vIVA_Porc = CCur(rs!Valor) / 100
  End Select
  
  rs.MoveNext
 Loop
 rs.Close
 
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
 
 vPaso = False
 
 Call sbLimpiaPantalla
 
 Call cboUnidad_Click
 Call cboDivisa_Click
  
 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = ""

txtFactura.Text = ""
txtFacturaNew.Text = ""
txtPlantilla = ""

imgEstado.Visible = False
imgCambio.Visible = False

gbFactura.Visible = False

txtEstado.Text = ""

dtpFecha.Value = fxFechaServidor
txtFecha = Format(dtpFecha.Value, "yyyy/mm/dd hh:mm:ss")
dtpFecha.Visible = True
dtpVence.Value = dtpFecha.Value

txtNotas = ""

vGrid.MaxRows = 2
vGrid.MaxCols = 8
For i = 1 To vGrid.MaxCols
  vGrid.Col = i
  vGrid.Text = ""
Next

cboDivisa.Text = vDivisaLocal
txtTipoCambio.Text = 1
txtTotalPagar.Text = 0
txtDivisaLocal.Text = 0
txtImpuesto.Text = 0

cboTipo.Clear
cboTipo.AddItem "Contado"
cboTipo.AddItem "Crédito"
cboTipo.Text = "Crédito"


StatusBarX.Panels.Item(1).Text = "Registrado por:"
StatusBarX.Panels.Item(1).ToolTipText = ""
StatusBarX.Panels.Item(2).Text = "Saldo: 0.00"


End Sub

Private Sub sbSumaDebitosCreditos()
Dim x As Integer, TC As Currency
  
On Error GoTo vError
  
  txtDebito = 0
  txtCredito = 0
  For x = 1 To vGrid.MaxRows
     vGrid.Row = x
     
     vGrid.Col = 5
     TC = 1 'Siempre es uno
      
     vGrid.Col = 7
     txtDebito = CCur(txtDebito) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * TC)
     vGrid.Col = 8
     txtCredito = CCur(txtCredito) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * TC)
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

vError:

End Sub



Private Sub sbCargaAsiento()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass
vGrid.MaxCols = 8


If Not IsNumeric(txtTotalPagar.Text) Then
  vGrid.MaxRows = 0
  Me.MousePointer = vbDefault
  Exit Sub
End If

'If CCur(txtTotalPagar.Text) = 0 Then
'  vGrid.MaxRows = 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'End If


If vCodigo <> "" Then
   strSQL = "select C.cod_Cuenta_Mask, C.cod_cuenta,C.descripcion as 'Cuenta',D.debehaber,D.monto,D.cod_unidad" _
          & ",U.descripcion as 'Unidad',D.cod_centro_costo,X.descripcion as 'CentroCosto',D.cod_proveedor,D.cod_factura" _
          & ",Div.Cod_Divisa,Div.Descripcion as 'Divisa',D.Tipo_Cambio" _
          & " from CXP_FACTURAS_DETALLE D inner join CXP_FACTURAS Ch on D.cod_factura = Ch.cod_factura and D.cod_proveedor = Ch.Cod_Proveedor" _
          & " inner join CntX_Cuentas C on D.cod_cuenta = C.cod_cuenta and D.cod_contabilidad = C.cod_Contabilidad" _
          & " inner join CntX_Divisas Div on D.cod_divisa = Div.cod_Divisa and D.cod_contabilidad = Div.cod_Contabilidad" _
          & "  left join CntX_unidades U on D.cod_unidad = U.cod_unidad and U.cod_contabilidad = D.cod_Contabilidad " _
          & "  left join CNTX_CENTRO_COSTOS X on D.cod_centro_costo = X.COD_CENTRO_COSTO and X.cod_contabilidad = " & GLOBALES.gEnlace _
          & " where D.cod_factura = '" & vCodigo & "' and D.cod_proveedor = " & txtProvCod.Text _
          & " order by D.linea"
          
    Call OpenRecordSet(rs, strSQL, 0)
    
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    
    
    Do While Not rs.EOF
      vGrid.Row = vGrid.MaxRows
      
      For i = 1 To vGrid.MaxCols
        vGrid.Col = i
        Select Case i
         Case 1 'Cuenta
             vGrid.Text = rs!Cod_Cuenta_Mask  'fxgCntCuentaFormato(True, CStr(rs!cod_cuenta))
         
         Case 2 'Unidad
            vGrid.Text = rs!Cod_Unidad & ""
            vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            vGrid.CellNote = rs!Unidad & ""
            vGrid.TextTip = TextTipFixed
         
         
         Case 3 'Centro de Costo
            vGrid.Text = rs!Cod_Centro_Costo & ""
            vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            vGrid.CellNote = rs!CentroCosto & ""
            vGrid.TextTip = TextTipFixed
         
         Case 4 'Divisa
            vGrid.Text = rs!COD_DIVISA & ""
            vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            vGrid.CellNote = rs!Divisa & ""
            vGrid.TextTip = TextTipFixed
         
         Case 5 'Tipo de Cambio
            vGrid.Text = CStr(rs!TIPO_CAMBIO)
         
         Case 6 'Descripcion
            vGrid.Text = CStr(rs!Cuenta)
         
         Case 7 'Debitos
           If rs!debehaber = "D" Then
             vGrid.Text = CStr(rs!Monto)
           Else
             vGrid.Text = "0"
           End If
         Case 8 'Creditos
           If rs!debehaber = "D" Then
             vGrid.Text = "0"
           Else
             vGrid.Text = CStr(rs!Monto)
           End If
        End Select
      Next i
      vGrid.MaxRows = vGrid.MaxRows + 1
      
      rs.MoveNext
    Loop
    rs.Close
    vGrid.MaxRows = vGrid.MaxRows - 1

Else 'vCodigo > 0
  
  vGrid.MaxRows = 0
  vGrid.MaxRows = 2
  vGrid.Row = 1
    
    
    strSQL = "select cod_cuenta from cxp_proveedores where cod_proveedor = " & txtProvCod.Text
    Call OpenRecordSet(rs, strSQL)
     vGrid.Col = 1
     vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
     vGrid.Col = 6
     vGrid.Text = fxgCntCuentaDesc(rs!cod_cuenta)
    rs.Close
    
    vGrid.Col = 2
    vGrid.Text = cboUnidad.ItemData(cboUnidad.ListIndex)
    strSQL = "select descripcion from CntX_Unidades where cod_unidad = '" & vGrid.Text & "' and cod_contabilidad = " & GLOBALES.gEnlace
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
     vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
     vGrid.CellNote = rs!Descripcion
     vGrid.TextTip = TextTipFixed
    End If
    rs.Close
     
     
    vGrid.Col = 3
    vGrid.Text = SIFGlobal.fxCodText(cboCentroCosto.Text)
    strSQL = "select descripcion from CntX_Centro_Costos where cod_centro_costo = '" & vGrid.Text & "' and cod_contabilidad = " & GLOBALES.gEnlace
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
     vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
     vGrid.CellNote = rs!Descripcion
     vGrid.TextTip = TextTipFixed
    End If
    rs.Close
     
    vGrid.Col = 4
    vGrid.Text = txtProvDivisa.Text
    strSQL = "select descripcion from CntX_Divisas where cod_divisa = '" & vGrid.Text & "' and cod_contabilidad = " & GLOBALES.gEnlace
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
     vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
     vGrid.CellNote = rs!Descripcion
     vGrid.TextTip = TextTipFixed
    End If
    rs.Close
     
    vGrid.Col = 5
    If vDivisa = Trim(txtProvDivisa.Text) Then
      vGrid.Text = "1"
    Else
      vGrid.Text = txtTipoCambio.Text
    End If
     
    vGrid.Row = 1
    vGrid.Col = 7
    vGrid.Text = "0"
    vGrid.Col = 8
    vGrid.Text = CStr(CCur(txtDivisaLocal.Text))
    
    vGrid.Row = 2
    vGrid.Col = 2
    vGrid.Text = cboUnidad.ItemData(cboUnidad.ListIndex)
    vGrid.Col = 3
    vGrid.Text = SIFGlobal.fxCodText(cboCentroCosto.Text)
    vGrid.Col = 7
    vGrid.Text = CStr(CCur(txtDivisaLocal.Text))
    vGrid.Col = 8
    vGrid.Text = "0"


End If 'vCodigo > 0


Call sbSumaDebitosCreditos

'Bloquea la Primer Linea
vGrid.Row = 1
For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Lock = True
    vGrid.Protect = True
Next i

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub sbPlantillaAsiento()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass
vGrid.MaxCols = 8


If Not IsNumeric(txtTotalPagar.Text) Then
  vGrid.MaxRows = 0
  Me.MousePointer = vbDefault
  Exit Sub
End If


strSQL = "select Cta.COD_CUENTA_MASK, Cta.DESCRIPCION, P.COD_UNIDAD, P.COD_CENTRO_COSTO, Cta.COD_DIVISA" _
        & "        , dbo.fxCntXTipoCambio(P.COD_CONTABILIDAD, Cta.COD_DIVISA, '" & Format(dtpFecha.Value, "yyyy-mm-dd") & "', 'V') as 'Tipo_Cambio'" _
        & "        , " & CCur(txtDivisaLocal.Text) & " * P.PORCENTAJE / 100 as 'Debito', 0 as 'Credito'" _
        & "        , isnull(D.DESCRIPCION,'') as 'Divisa_Desc'" _
        & "        , isnull(U.DESCRIPCION,'') as 'Unidad_Desc', isnull(C.DESCRIPCION,'') as 'Centro_Desc'" _
        & " from CXP_PLANTILLAS_ASIENTO P inner join CNTX_CUENTAS Cta on P.COD_CONTABILIDAD = Cta.COD_CONTABILIDAD" _
        & "    and P.COD_CUENTA = Cta.COD_CUENTA" _
        & "        left join CNTX_DIVISAS D on Cta.COD_CONTABILIDAD = D.COD_CONTABILIDAD and   Cta.COD_DIVISA = D.COD_DIVISA" _
        & "        left join CNTX_UNIDADES U on P.COD_CONTABILIDAD = U.COD_CONTABILIDAD and P.COD_UNIDAD = U.COD_UNIDAD" _
        & "        left join CNTX_CENTRO_COSTOS  C on P.COD_CONTABILIDAD = C.COD_CONTABILIDAD and P.COD_CENTRO_COSTO = C.COD_CENTRO_COSTO" _
        & " Where COD_PLANTILLA = '" & txtPlantilla.Text & "'" _
        & " order by LINEA"

       
 Call OpenRecordSet(rs, strSQL, 0)
 
 vGrid.MaxRows = 2
 Do While Not rs.EOF
   vGrid.Row = vGrid.MaxRows
   
   For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     Select Case i
      Case 1 'Cuenta
          vGrid.Text = rs!Cod_Cuenta_Mask
      
      Case 2 'Unidad
         vGrid.Text = rs!Cod_Unidad & ""
         vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
         vGrid.CellNote = rs!Unidad_Desc & ""
         vGrid.TextTip = TextTipFixed
      
      
      Case 3 'Centro de Costo
         vGrid.Text = rs!Cod_Centro_Costo & ""
         vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
         vGrid.CellNote = rs!Centro_Desc & ""
         vGrid.TextTip = TextTipFixed
      
      Case 4 'Divisa
         vGrid.Text = rs!COD_DIVISA & ""
         vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
         vGrid.CellNote = rs!Divisa_Desc & ""
         vGrid.TextTip = TextTipFixed
      
      Case 5 'Tipo de Cambio
         vGrid.Text = CStr(rs!TIPO_CAMBIO)
      
      Case 6 'Descripcion
         vGrid.Text = CStr(rs!Descripcion)
      
      Case 7 'Debitos
          vGrid.Text = CStr(rs!Debito)
      Case 8 'Creditos
          vGrid.Text = "0"
     End Select
   Next i
   vGrid.MaxRows = vGrid.MaxRows + 1
   
   rs.MoveNext
 Loop
 rs.Close
 vGrid.MaxRows = vGrid.MaxRows - 1


Call sbSumaDebitosCreditos

'Bloquea la Primer Linea
vGrid.Row = 1
For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Lock = True
    vGrid.Protect = True
Next i

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




Private Sub imgCambio_Click()
If gbFactura.Visible Then
    gbFactura.Visible = False
Else
    gbFactura.Visible = True
    gbFactura.Left = 120
    gbFactura.top = 960
    
End If
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtFactura.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNotas.SetFocus
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
'       gBusquedas.Columna = "descripcion"
'       gBusquedas.Orden = "descripcion"
'       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'       frmBusquedas.Show vbModal
'       txtFactura.SetFocus
'       txtFactura = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
'       txtNombre.SetFocus

    Case "REPORTES"
    
     Call sbReportes

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub sbReportes()

If Not IsNumeric(txtProvCod.Text) Then Exit Sub

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Cuentas por Pagar"

    .Connect = glogon.ConectRPT
    .Formulas(1) = "fxUsuario = '" & glogon.Usuario & "'"
    .Formulas(2) = "fxFecha = '" & fxFechaServidor & "'"
    .Formulas(3) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("CxP_Factura_Registro.rpt")
    .SelectionFormula = "{vCxP_Facturas_Main.COD_PROVEEDOR} = " & txtProvCod.Text _
                      & " AND {vCxP_Facturas_Main.COD_FACTURA} = '" & txtFactura.Text & "'"

    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub



Private Sub sbConsulta(pFactura As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select F.*,P.descripcion as Proveedor,P.cod_Divisa as 'DivisaProv'" _
       & ",dbo.fxCxPSaldoFacturaCorte(F.cod_Proveedor,F.cod_Factura,dbo.MyGetdate()) as 'Saldo'" _
       & ",rtrim(D.descripcion) as 'DivisaFactura'" _
       & " from cxp_facturas F inner join cxp_proveedores P on F.cod_proveedor = P.cod_proveedor" _
       & " inner join CntX_Divisas D on D.cod_contabilidad = " & GLOBALES.gEnlace & " and D.cod_divisa = F.cod_divisa" _
       & " where F.cod_factura = '" & pFactura & "'"
If txtProvCod <> "" Then
   strSQL = strSQL & " and F.cod_proveedor = " & txtProvCod
End If
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  

  vCodigo = rs!cod_Factura
  txtFactura = rs!cod_Factura
  txtFacturaNew = rs!cod_Factura
  
  imgEstado.Visible = True
  imgCambio.Visible = True
  gbFactura.Visible = False
  
  StatusBarX.Panels(1).ToolTipText = ""
  StatusBarX.Panels(2).Text = "Saldo: " & Format(rs!Saldo, "Standard")
  
  vPaso = True
    Call sbCboAsignaDato(cboDivisa, rs!DivisaFactura, True, rs!COD_DIVISA)
  vPaso = False
  
  Select Case UCase(Trim(rs!Estado))
    Case "P"
         txtEstado.Text = "Procesada"
         StatusBarX.Panels(1).Text = "Registrado por: " & Trim(rs!creacion_user) & " - " & rs!Creacion_Fecha
         If rs!CxP_Estado = "P" Then
            imgEstado.Picture = ImageList1.ListImages.Item(6).Picture
            imgEstado.ToolTipText = "Factura Activa! pero Programada"
         Else
            imgEstado.Picture = ImageList1.ListImages.Item(5).Picture
            imgEstado.ToolTipText = "Factura Activa! y Programada"
         End If
    Case "A"
            txtEstado.Text = "Anulada"
            imgEstado.Picture = ImageList1.ListImages.Item(7).Picture
            imgEstado.ToolTipText = "Factura Anulada!"
            StatusBarX.Panels(1).Text = "Anulada por: " & rs!anula_user & " - " & rs!anula_fecha
            StatusBarX.Panels(1).ToolTipText = "Registrado por: " & Trim(rs!creacion_user) & " - " & rs!Creacion_Fecha
    
    Case Else
         txtEstado.Text = "Estado!"
  End Select
  
  txtProvCod = rs!cod_Proveedor
  txtProvDesc = rs!Proveedor
  
  txtFecha = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  dtpFecha.Value = rs!fecha
  
  dtpVence.Value = rs!Vence
  txtNotas = rs!Notas & ""
  
  txtTipoCambio.Text = Format(rs!TIPO_CAMBIO, "###,###.00####")
  
  If rs!DivisaFactura <> vDivisa Then
        StatusBarX.Panels(2).Text = "Saldo: " & Format(rs!Saldo, "Standard") & Space(10) & "[ " & Format(rs!Saldo / rs!TIPO_CAMBIO, "Standard") & " ]"
        txtTotalPagar.Text = Format(rs!Importe_divisa_real, "Standard")
        txtDivisaLocal.Text = Format(rs!Total, "Standard")
  
  Else
        txtTotalPagar.Text = Format(rs!Total, "Standard")
        txtDivisaLocal.Text = Format(rs!Importe_divisa_real, "Standard")
  End If
  
  
  txtProvDivisa.Text = rs!DivisaProv
  
  chkPlantilla.Value = rs!plantilla
  
  Select Case UCase(Trim(rs!Cod_Forma_Pago))
    Case "CO"
        cboTipo.Text = "Contado"
    Case "CR"
        cboTipo.Text = "Crédito"
  End Select
  
  txtImpuesto.Text = Format(rs!impuesto_ventas, "Standard")
  Call sbCargaAsiento
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
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
Dim vMensaje As String, vTemp As String
Dim curMonto As Currency, i As Integer

vMensaje = ""
fxValida = True


'Limpieza de Inyeccion

txtFactura.Text = fxSysCleanTxtInject(txtFactura.Text)
txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)


'If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."
'If dtpFecha.Visible Then
'    If Not fxInvPeriodos(dtpFecha.Value) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."
'Else
'    If Not fxInvPeriodos(fxFechaServidor) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."
'End If

If IsNumeric(txtTotalPagar.Text) Then
   If CCur(txtTotalPagar) <= 0 Then
        vMensaje = vMensaje & vbCrLf & " - El monto de la Factura no es válido..."
   End If
Else
    vMensaje = vMensaje & vbCrLf & " - El monto de la Factura no es válido..."
End If


If IsNumeric(txtImpuesto.Text) Then
   If CCur(txtImpuesto.Text) < 0 Then
        vMensaje = vMensaje & vbCrLf & " - El monto del IMPUESTO no es válido..."
   End If
Else
    vMensaje = vMensaje & vbCrLf & " - El monto del IMPUESTO no es válido..."
End If

If Trim(txtFactura.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Número de factura no es válida ..."
If txtProvCod.Text = "" Or Not IsNumeric(txtProvCod.Text) Then vMensaje = vMensaje & vbCrLf & " - Código del Proveedor no es válido ..."
If Trim(txtProvDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Proveedor no es válido ..."

'Si la divisa del Proveedor no es igual que la divisa local, validar que la factura sea solo en su divisa origen
If vDivisa <> Trim(txtProvDivisa.Text) Then
    vTemp = cboDivisa.ItemData(cboDivisa.ListIndex)
    If vTemp <> Trim(txtProvDivisa.Text) Then
       vMensaje = vMensaje & vbCrLf & " - La divisa utilizada en la factura no es válida (No concuerda con la del proveedor)..."
    End If
End If

'Revisar Asiento, si no tiene Lineas crear Asiento Básico
If vGrid.MaxRows <= 1 Then Call sbCargaAsiento

'despues de la creacion del asiento revisar #lineas y Balance
If vGrid.MaxRows < 2 Then vMensaje = vMensaje & vbCrLf & " - El Asiento no se válido..."

Call sbSumaDebitosCreditos
If CCur(txtDiferencia) <> 0 Then vMensaje = vMensaje & vbCrLf & " - El Asiento no se encuentra balanceado..."

'Valida que la Primer linea del Asiento sea igual al monto del documento
vGrid.Row = 1
vGrid.Col = 7
curMonto = CCur(vGrid.Text)
vGrid.Col = 8
curMonto = curMonto + CCur(vGrid.Text)
If curMonto <> CCur(txtDivisaLocal.Text) Then vMensaje = vMensaje & vbCrLf & " - El Monto Linea 1 del Asiento no corresponde al original..."


'Valida Cuenta, Unidad de Negocios y Centro de Costo
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 1
  If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, vGrid.Text)) Then
      vMensaje = vMensaje & vbCrLf & " - La cuenta de la linea : " & i & " no es válida..."
  End If
  
  vGrid.Col = 2
  If fxgCntUnidad(vGrid.Text) = "" Then
     vMensaje = vMensaje & vbCrLf & " - La unidad de negocios no es válida en la línea : " & i
  End If

  vGrid.Col = 3
  If fxgCntCentroCostos(vGrid.Text) = "" And vGrid.Text <> "" Then
     vMensaje = vMensaje & vbCrLf & " - El Centro de Costo no es válido en la línea : " & i
  End If
    
Next i

vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If


End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer
Dim vFecha As Date, curCantidad As Currency
Dim vTipoCambio As Currency

On Error GoTo vError

If vEdita Then
   MsgBox "No se puede editar una factura Guardada...", vbInformation
   Exit Sub
End If

If dtpFecha.Visible Then
  vFecha = dtpFecha.Value
Else
  vFecha = fxFechaServidor
End If


txtFactura.Text = Trim(txtFactura.Text)
vCodigo = txtFactura.Text

strSQL = "insert cxp_facturas(estado,cod_factura,cod_proveedor,fecha,total,cxp_estado" _
       & ",asiento_generado,plantilla,vence,creacion_fecha,creacion_user,notas,cod_forma_Pago" _
       & ",cod_divisa,tipo_cambio,importe_divisa_real,IMPUESTO_VENTAS)" _
       & " values('P','" & txtFactura & "'," & txtProvCod & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") _
       & "'," & CCur(txtDivisaLocal.Text) & ",'" & IIf((UCase(Mid(cboTipo.Text, 1, 2)) = "CR"), "P", "G") _
       & "','P'," & chkPlantilla.Value & ",'" & Format(dtpVence.Value, "yyyy/mm/dd") _
       & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtNotas & "','" & UCase(Mid(cboTipo.Text, 1, 2)) _
       & "','" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & CCur(txtTipoCambio.Text) _
       & "," & CCur(txtTotalPagar.Text) & "," & CCur(txtImpuesto.Text) & ")"
Call ConectionExecute(strSQL)

'Actualiza Saldo de la Cuenta por Pagar al Proveedor
'Supuesto: La divisa de la factura esta validad para que sea en la divisa local o en la del proveedor
If vDivisa = Trim(txtProvDivisa.Text) Then
    strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) + " & CCur(txtDivisaLocal.Text) _
           & ",SALDO_DIVISA_REAL =  isnull(SALDO_DIVISA_REAL ,0) + " & CCur(txtDivisaLocal.Text) _
           & " where cod_proveedor = " & txtProvCod
Else
    strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) + " & CCur(txtDivisaLocal.Text) _
           & ",SALDO_DIVISA_REAL =  isnull(SALDO_DIVISA_REAL ,0) + " & CCur(txtTotalPagar.Text) _
           & " where cod_proveedor = " & txtProvCod
End If
Call ConectionExecute(strSQL)


'Registrar Pagos al Contado dentro del programacion de pagos
If UCase(Mid(cboTipo.Text, 1, 2)) = "CO" Then
  strSQL = "insert cxp_pagoProv(NPago,Cod_Proveedor,Cod_Factura,Fecha_Vencimiento,Monto,Frecuencia" _
         & ",Tipo_Transac,User_TrasLada,Fecha_Traslada,Tesoreria,Pago_Tercero,Apl_Cargo_Flotante" _
         & ",Pago_Anticipado,forma_pago,IMPORTE_DIVISA_REAL,TIPO_CAMBIO,COD_DIVISA)" _
         & " values(1," & txtProvCod & ",'" & txtFactura & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "'," & CCur(txtDivisaLocal.Text) _
         & ",0,0,Null,Null,Null,''," & chkCargosFlotantesAplica.Value & ",0,'CO'," & CCur(txtTotalPagar.Text) & "," & CCur(txtTipoCambio.Text) _
         & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "')"
  Call ConectionExecute(strSQL)
End If

Call Bitacora("Registra", "CxP-Factura: " & vCodigo & "...Prov:" & txtProvCod.Text)

'Inicia Batch para el Detalle de la Factura
strSQL = "delete cxp_facturas_detalle" _
         & " where cod_factura = '" & txtFactura & "' and cod_proveedor = " & txtProvCod

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 1
  If vGrid.Text <> "" Then
    strSQL = strSQL & Space(10) & "insert cxp_facturas_detalle(linea,cod_factura,cod_proveedor,cod_contabilidad,cod_cuenta,cod_unidad,cod_centro_costo,cod_divisa" _
                                       & ",debeHaber,tipo_cambio,Monto) values(" & i & ",'" & txtFactura & "'," & txtProvCod & "," _
                                       & GLOBALES.gEnlace & ",'" & fxgCntCuentaFormato(False, vGrid.Text) & "','"
    vGrid.Col = 2
    strSQL = strSQL & Trim(vGrid.Text) & "','"
    vGrid.Col = 3
    strSQL = strSQL & Trim(vGrid.Text) & "','"
    vGrid.Col = 4
    strSQL = strSQL & Trim(vGrid.Text) & "','"
    vGrid.Col = 5
    vTipoCambio = CCur(vGrid.Text)
    
    vGrid.Col = 7
    If Not IsNumeric(vGrid.Text) Then vGrid.Text = "0"
    If CCur(vGrid.Text) > 0 Then
       strSQL = strSQL & "D'," & vTipoCambio & "," & CCur(vGrid.Text) & ")"
    Else
        vGrid.Col = 8
        If Not IsNumeric(vGrid.Text) Then vGrid.Text = "0"
        strSQL = strSQL & "H'," & vTipoCambio & "," & CCur(vGrid.Text) & ")"
    End If
    
  
  End If

Next i

'Aplica el Batch del detalle de la factura
Call ConectionExecute(strSQL)


MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(txtFactura.Text)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
'  Call ConectionExecute(strSQL)

'  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & vParametros.CodigoEmpresa)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAnular()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCxPFacturaAnula " & txtProvCod.Text & ",'" & Trim(txtFactura.Text) & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbConsulta(Trim(txtFactura.Text))

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvCod.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_factura"
  gBusquedas.Orden = "cod_factura"
  gBusquedas.Consulta = "select cod_factura,cod_proveedor,TOTAL,notas from cxp_facturas"
  
  gBusquedas.Filtro = ""
  If txtProvCod.Text <> "" Then
      gBusquedas.Filtro = " and cod_proveedor = " & txtProvCod.Text
  End If
  
  frmBusquedas.Show vbModal
  txtFactura = gBusquedas.Resultado
  If txtFactura <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtFactura_LostFocus()
If txtFactura <> "" And vEdita Then Call sbConsulta(txtFactura)
End Sub



Private Sub txtImpuesto_GotFocus()
On Error GoTo vError
    txtImpuesto.Text = CCur(txtImpuesto)
Exit Sub
vError:
End Sub

Private Sub txtImpuesto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 cboTipo.SetFocus
End If
End Sub

Private Sub txtImpuesto_LostFocus()
On Error GoTo vError
    txtImpuesto.Text = Format(CCur(txtImpuesto.Text), "Standard")

    Call sbCalculoDivisaLocal
    Call sbCambiaInfo

Exit Sub
vError:
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotalPagar.SetFocus
End Sub



Private Sub txtPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "COD_PLANTILLA"
  gBusquedas.Orden = "COD_PLANTILLA"
  gBusquedas.Consulta = "select COD_PLANTILLA, DESCRIPCION  From CXP_PLANTILLAS"
  gBusquedas.Filtro = " and ACTIVO = 1"
  frmBusquedas.Show vbModal
  txtPlantilla.Text = gBusquedas.Resultado
End If
End Sub

Private Sub txtProvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion, cedjur from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtProvCod.Text) Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select P.cod_proveedor,P.descripcion,P.cod_divisa" _
       & ",rtrim(D.descripcion) as 'DivisaLocal'" _
       & " from  Cxp_Proveedores P inner join CntX_Divisas D on P.cod_divisa = D.cod_divisa" _
       & " and D.cod_contabilidad = " & GLOBALES.gEnlace _
       & " where P.cod_proveedor = " & txtProvCod.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtProvDesc.Text = rs!Descripcion
  txtProvDivisa.Text = rs!COD_DIVISA
  
  Call sbCboAsignaDato(cboDivisa, rs!DivisaLocal, True, rs!COD_DIVISA)
Else
  txtProvCod.Text = ""
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub txtProvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If

End Sub


Private Function fxCuentaProveedor(pCodPro As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As String

strSQL = "select cod_cuenta from cxp_proveedores where cod_proveedor = " & pCodPro
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  vResultado = ""
Else
  vResultado = Trim(rs!cod_cuenta)
End If
rs.Close

fxCuentaProveedor = vResultado

End Function


Private Sub sbCambiaInfo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, vDescripcion As String, i As Integer

If Not IsNumeric(txtTotalPagar) Or txtProvCod.Text = "" Then Exit Sub
If CCur(txtTotalPagar.Text) = 0 Then Exit Sub

strSQL = "select C.cod_cuenta,C.descripcion,P.cod_Divisa as 'DivisaProv'" _
       & " from cxp_proveedores P inner join Cntx_Cuentas C on P.cod_cuenta = C.cod_cuenta and C.cod_Contabilidad = " & GLOBALES.gEnlace _
       & " where P.cod_proveedor = " & txtProvCod.Text

Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  vCuenta = ""
  vDescripcion = ""
Else
  vCuenta = Trim(rs!cod_cuenta)
  vDescripcion = Trim(rs!Descripcion)
  txtProvDivisa.Text = Trim(rs!DivisaProv)
End If
rs.Close

With vGrid
   .Row = 1
   .Col = 1
   .Text = fxgCntCuentaFormato(True, vCuenta)
   .Col = 6
   .Text = vDescripcion
   .Col = 2
   .Text = cboUnidad.ItemData(cboUnidad.ListIndex)
   .Col = 3
   .Text = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
   
   .Col = 4 'Divisa
   .Text = txtProvDivisa.Text
   
   .Col = 5 'Tipo de Cambio
   
   If txtProvDivisa.Text = vDivisa Then
        .Text = "1"
   Else
        .Text = txtTipoCambio.Text
   End If
   
   .Col = 7
   .Text = 0
   .Col = 8
   .Text = CCur(txtDivisaLocal.Text)
   
   '-------------------------------------------------------------------------------------
   'Registra la Linea del IVA
   .Row = 2
   .Col = 1
   .Text = fxgCntCuentaFormato(True, vIVA_Cta)
   .Col = 6
   .Text = vIVA_CtaDesc
   .Col = 2
   .Text = cboUnidad.ItemData(cboUnidad.ListIndex)
   .Col = 3
   .Text = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
   
   .Col = 4 'Divisa
   .Text = txtProvDivisa.Text
   
   .Col = 5 'Tipo de Cambio
   If txtProvDivisa.Text = vDivisa Then
        .Text = "1"
   Else
        .Text = txtTipoCambio.Text
   End If
   
   .Col = 7
   .Text = CCur(txtImpuesto.Text) * fxSys_Tipo_Cambio_Apl(CCur(txtTipoCambio.Text))
   .Col = 8
   .Text = 0
   
End With

'Bloquea la Primer Linea
vGrid.Row = 1
For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Lock = True
    vGrid.Protect = True
Next i


Call sbSumaDebitosCreditos

End Sub

Private Sub txtTotalPagar_GotFocus()
On Error GoTo vError
    txtTotalPagar.Text = CCur(txtTotalPagar)
Exit Sub
vError:
End Sub

Private Sub txtTotalPagar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 cboTipo.SetFocus
End If
End Sub


Private Sub txtTotalPagar_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

txtImpuesto.Text = CCur(txtTotalPagar.Text) - (CCur(txtTotalPagar.Text) / (1 + vIVA_Porc))
txtImpuesto.Text = Format(CCur(txtImpuesto.Text), "Standard")


vError:
End Sub

Private Sub txtTotalPagar_LostFocus()
On Error GoTo vError
    txtTotalPagar.Text = Format(CCur(txtTotalPagar.Text), "Standard")
    Call sbCalculoDivisaLocal
    Call sbCambiaInfo
Exit Sub
vError:
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer
Dim vCuenta As String

'No permite Borrar la Primer linea, las demás SI
If KeyCode = vbKeyDelete And vGrid.ActiveRow > 1 Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = vGrid.MaxCols
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  
  
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.Col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
End If

'Consulta cuenta / Codigo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 And vGrid.ActiveRow > 1 Then
  Call sbgCntCuentaConsulta("C")
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

'Consulta cuenta / descripcion
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 6 And vGrid.ActiveRow > 1 Then
  Call sbgCntCuentaConsulta("D")
  vGrid.Col = 1
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If


'Consulta unidad
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 2 And vGrid.ActiveRow > 1 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and Activa = 1 and cod_contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Consulta = "select cod_unidad,descripcion from CntX_unidades"
  frmBusquedas.Show vbModal
    
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
  
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  
End If



'Consulta Centro de Costo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 And vGrid.ActiveRow > 1 Then
  
  vGrid.Col = 2
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = " and C.cod_Contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select C.COD_CENTRO_COSTO,C.descripcion" _
                      & " from CNTX_CENTRO_COSTOS C inner join CNTX_UNIDADES_CC A on C.COD_CENTRO_COSTO = A.COD_CENTRO_COSTO" _
                      & " and C.cod_contabilidad = A.cod_Contabilidad" _
                      & " and A.cod_unidad = '" & vGrid.Text & "'"
  frmBusquedas.Show vbModal
    
  vGrid.Col = 3
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
  
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  
End If



If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text, 0)
        
        If fxgCntCuentaValida(fxgCntCuentaFormato(False, vGrid.Text, 0)) Then
            vCuenta = vGrid.Text
            vCuenta = fxgCntCuentaFormato(False, vCuenta, 0)
'            vGrid.Col = 6
'
'            vGrid.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, vCuenta, 0))

            strSQL = "select Descripcion,COD_DIVISA" _
                   & ",dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",COD_DIVISA,dbo.MyGetdate(),'V') as 'Tipo_Cambio'" _
                   & " from CntX_Cuentas where cod_cuenta = '" & vCuenta & "' and cod_contabilidad = " & GLOBALES.gEnlace
            Call OpenRecordSet(rs, strSQL)
            If Not rs.BOF And Not rs.EOF Then
                vGrid.Col = 6
                vGrid.Text = Trim(rs!Descripcion)
                vGrid.Col = 4
                vGrid.Text = Trim(rs!COD_DIVISA)
                vGrid.Col = 5
                vGrid.Text = CStr(rs!TIPO_CAMBIO)
            
            End If
            rs.Close

        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CUENTAS", vbCritical
        End If
        
      Case 2
        'Buscar la unidad
        If fxgCntUnidad(vGrid.Text) <> "" Then
          vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
          vGrid.CellNote = fxgCntUnidad(vGrid.Text)
          vGrid.TextTip = TextTipFixed
        Else
          MsgBox "La unidad de negocio no es válida, o no se tiene asignada a este usuario", vbCritical
        End If
      
      Case 3 'Describe el centro de costo
          vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
          vGrid.CellNote = fxgCntCentroCostos(vGrid.Text)
          vGrid.TextTip = TextTipFixed
      
        
      Case 7 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 8 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
        End If
      
        If vGrid.MaxRows = vGrid.Row Then
            
            vGrid.Col = 2
            vTemp(0) = vGrid.Text
            vTemp(1) = vGrid.CellNote
            
            vGrid.Col = 3
            vTemp(2) = vGrid.Text
            vTemp(3) = vGrid.CellNote
            
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.MaxRows
        
            vGrid.Col = 2
            vGrid.Text = vTemp(0)
            vGrid.CellNote = vTemp(1)
            
            vGrid.Col = 3
            vGrid.Text = vTemp(2)
            vGrid.CellNote = vTemp(3)
        End If
    
    End Select

End If

If KeyCode = vbKeyInsert And vGrid.ActiveRow > 1 Then
    vGrid.Row = vGrid.ActiveRow
            vGrid.Col = 2
            vTemp(0) = vGrid.Text
            vTemp(1) = vGrid.CellNote
            
            vGrid.Col = 3
            vTemp(2) = vGrid.Text
            vTemp(3) = vGrid.CellNote
    
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
            
            vGrid.Col = 2
            vGrid.Text = vTemp(0)
            vGrid.CellNote = vTemp(1)
            
            vGrid.Col = 3
            vGrid.Text = vTemp(2)
            vGrid.CellNote = vTemp(3)
    vGrid.Col = 1
End If


End Sub

