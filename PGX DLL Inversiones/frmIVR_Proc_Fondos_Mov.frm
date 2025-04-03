VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmIVR_Proc_Fondos_Mov 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Movimientos a Fondos de Inversión"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lswAd 
      Height          =   1812
      Left            =   0
      TabIndex        =   16
      Top             =   6000
      Width           =   11172
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   3196
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
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   3840
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox gbFondos 
      Height          =   2772
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   11052
      _Version        =   1441793
      _ExtentX        =   19494
      _ExtentY        =   4890
      _StockProps     =   79
      Caption         =   "Fondos de Inversion:"
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
      BorderStyle     =   2
      Begin XtremeSuiteControls.DateTimePicker dtpFi_FechaMovimiento 
         Height          =   312
         Left            =   3120
         TabIndex        =   1
         Top             =   0
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtFi_ValorActual 
         Height          =   312
         Left            =   7200
         TabIndex        =   24
         Top             =   1680
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   552
         Left            =   3120
         TabIndex        =   27
         Top             =   2160
         Width           =   7452
         _Version        =   1441793
         _ExtentX        =   13144
         _ExtentY        =   974
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
      Begin XtremeSuiteControls.FlatEdit txtFi_ParticipacionNo 
         Height          =   312
         Left            =   3120
         TabIndex        =   5
         Top             =   1680
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFi_ParticipacionValor 
         Height          =   312
         Left            =   3120
         TabIndex        =   4
         Top             =   1320
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFi_Movimiento 
         Height          =   312
         Left            =   3120
         TabIndex        =   3
         Top             =   840
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   5160
         TabIndex        =   36
         Top             =   840
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
         Height          =   312
         Left            =   6000
         TabIndex        =   37
         Top             =   840
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtImporteLocal 
         Height          =   312
         Left            =   7200
         TabIndex        =   40
         Top             =   840
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   11
         Left            =   1080
         TabIndex        =   44
         Top             =   480
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "No. Documento "
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   10
         Left            =   7200
         TabIndex        =   41
         Top             =   600
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Importe Local:"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   7
         Left            =   6120
         TabIndex        =   39
         Top             =   600
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "T.C:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   5
         Left            =   5160
         TabIndex        =   38
         Top             =   600
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Divisa:"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   41
         Left            =   1080
         TabIndex        =   35
         Top             =   1680
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "No. de Participaciones"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   42
         Left            =   1080
         TabIndex        =   34
         Top             =   1320
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor de la Participación"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   44
         Left            =   1080
         TabIndex        =   33
         Top             =   840
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Monto del Movimiento"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   8
         Left            =   1080
         TabIndex        =   26
         Top             =   2160
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   6
         Left            =   7080
         TabIndex        =   25
         Top             =   1440
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor Actual del Fondo:"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   40
         Left            =   1080
         TabIndex        =   6
         Top             =   0
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha Movimiento"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtInversionId 
      Height          =   492
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtInstrumento 
      Height          =   312
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   9252
      _Version        =   1441793
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtAdministrador 
      Height          =   312
      Left            =   1800
      TabIndex        =   9
      Top             =   1200
      Width           =   9252
      _Version        =   1441793
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtPortafolio 
      Height          =   312
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   9252
      _Version        =   1441793
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.PushButton btnAdquisicion 
      Height          =   372
      Index           =   0
      Left            =   9000
      TabIndex        =   17
      Top             =   5580
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nuevo"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtAd_Pendiente 
      Height          =   312
      Left            =   5640
      TabIndex        =   18
      Top             =   8280
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAdquisicion 
      Height          =   372
      Index           =   1
      Left            =   10560
      TabIndex        =   19
      Top             =   5580
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
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
      Picture         =   "frmIVR_Proc_Fondos_Mov.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtAd_Registrado 
      Height          =   312
      Left            =   5640
      TabIndex        =   21
      Top             =   7920
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTransacId 
      Height          =   492
      Left            =   5760
      TabIndex        =   28
      Top             =   120
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   372
      Index           =   0
      Left            =   9000
      TabIndex        =   30
      Top             =   2100
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nueva Solicitud!"
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
   End
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   612
      Index           =   1
      Left            =   8280
      TabIndex        =   31
      Top             =   7920
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Procesar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmIVR_Proc_Fondos_Mov.frx":05A4
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   612
      Index           =   2
      Left            =   10080
      TabIndex        =   32
      Top             =   7920
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   1080
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
      Appearance      =   14
      Picture         =   "frmIVR_Proc_Fondos_Mov.frx":0F67
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   492
      Left            =   8640
      TabIndex        =   42
      Top             =   120
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin XtremeSuiteControls.FlatEdit txtSeqId 
      Height          =   492
      Left            =   7440
      TabIndex        =   43
      ToolTipText     =   "Consecutivo de Movimiento en el Fondo"
      Top             =   120
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   9
      Left            =   4200
      TabIndex        =   29
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Id Transac:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   492
      Left            =   0
      TabIndex        =   23
      Top             =   5520
      Width           =   11172
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Comprobantes:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   4
      Left            =   4440
      TabIndex        =   22
      Top             =   7920
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Registrado"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   35
      Left            =   4440
      TabIndex        =   20
      Top             =   8280
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pendiente"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Instrumento"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Administrador"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Portafolio"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Inversión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scGestion 
      Height          =   492
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   11172
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Gestion: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmIVR_Proc_Fondos_Mov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean
Dim itmX As ListViewItem, vFecha As Date
Dim vDivisaLocaL As String, pConsultaExternaId As Long


Private Sub sbAdquisicion_Load()
Call sbIVR_Transac_Load(lswAd, txtTransacId.Text, gIVR_Transito.Tipo, gIVR_Transito.Concepto)

Dim i As Integer, pMonto As Currency

With lswAd.ListItems

pMonto = 0
For i = 1 To .Count
    pMonto = pMonto + CCur(.Item(i).SubItems(3))
Next i

txtAd_Registrado.Text = Format(pMonto, "Standard")

If Not IsNumeric(txtImporteLocal.Text) Then
    txtImporteLocal.Text = "0"
End If

txtAd_Pendiente.Text = Format(CCur(txtImporteLocal.Text) - pMonto, "Standard")



End With
 
End Sub



Private Sub btnAdquisicion_Click(Index As Integer)



If Mid(txtEstado.Text, 1, 1) <> "S" Then
   Exit Sub
End If

Select Case Index
    Case 0 'Nuevo
        
        If CLng(txtTransacId.Text) = 0 Then
           Call sbGuardar
        End If
        
        
        gIVR_Transito.Codigo = txtTransacId.Text
        gIVR_Transito.Divisa = vDivisaLocaL
        gIVR_Transito.TipoCambio = 1
        gIVR_Transito.Monto = CCur(txtAd_Pendiente.Text)
        
        frmIVR_Rec_Bancos_Registro.Show vbModal
        
    Case 1 'Eliminar
        
        Dim i As Integer
        With lswAd.ListItems
            For i = 1 To .Count
                If .Item(i).Checked = True Then
                    strSQL = "delete  IVR_TRANSACCIONES Where TRANSAC_ID = " & .Item(i).Text
                    Call ConectionExecute(strSQL)
                End If
            Next i
        End With
End Select

Call sbAdquisicion_Load

End Sub


Private Sub btnTransac_Click(Index As Integer)

On Error GoTo vError

Select Case Index
    Case 0 'Nueva Solicitud
        Call sbInicializa
    
    Case 1 'Procesar
        If Len(txtDocumento.Text) = 0 Then
            MsgBox "No ha Indicado un No. de Documento para esta transacción!", vbExclamation
            Exit Sub
        End If
    
        If dtpFi_FechaMovimiento.Value < vFecha Then
            MsgBox "La fecha de esta transacción es menor al último corte!", vbExclamation
            Exit Sub
        End If
    
    
        If CCur(txtImporteLocal.Text) > 0 And CCur(txtAd_Pendiente.Text) = 0 Then
    
            If CLng(txtTransacId.Text) > 0 And Mid(txtEstado.Text, 1, 1) = "S" Then
               Call sbGuardar("A")
            End If
        
            MsgBox "Transacción procesada Satisfactoriamente!", vbInformation
        
        Else
            MsgBox "No se ha indicado ningún valor a la transacción!", vbExclamation
        End If
        
    Case 2 'Eliminar
        If CLng(txtTransacId.Text) > 0 And Mid(txtEstado.Text, 1, 1) = "S" Then
           Call sbGuardar("X")
           Call sbInicializa
        Else
            MsgBox "No se ha indicado una transacción y/o esta ya fue procesada!", vbExclamation
        End If
End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub



Private Sub dtpFi_FechaMovimiento_Change()

On Error GoTo vError

Me.MousePointer = vbHourglass

If txtDivisa.Text <> vDivisaLocaL Then
    strSQL = "select dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ", '" & txtDivisa.Text _
            & "', '" & Format(dtpFi_FechaMovimiento.Value, "yyyy/mm/dd") & "', 'V') as 'TipoCambio'"
    Call OpenRecordSet(rs, strSQL)
      txtTipoCambio.Text = Format(rs!TipoCambio, "###,##0.00000")
    rs.Close
End If

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

On Error GoTo vError

pConsultaExternaId = 0

txtInversionId.Text = gIVR_Transito.TituloId
txtTransacId.Text = "0"
txtEstado.Text = "Solicitud"

If gIVR_Transito.Concepto = "FI_APO" Then
     scGestion.Caption = "Gestión:   Aportación al Fondo"
Else
     scGestion.Caption = "Gestión:   Retiro del Fondo"
End If

strSQL = "select  isnull(max(CORTE), dbo.mygetdate())  as 'CORTE'" _
       & "  From IVR_CIERRES"
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!Corte
rs.Close

Exit Sub

vError:


End Sub



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


On Error GoTo vError

vPaso = True


strSQL = "select COD_DIVISA  From vSys_Divisas  Where DIVISA_LOCAL = 1"
Call OpenRecordSet(rs, strSQL)
   vDivisaLocaL = Trim(rs!Cod_Divisa)
rs.Close


txtDivisa.Text = gIVR_Transito.Divisa
txtTipoCambio.Text = gIVR_Transito.TipoCambio


If vDivisaLocaL = gIVR_Transito.Divisa Then
    txtTipoCambio.Text = "1"
    txtTipoCambio.Locked = True
Else
    txtTipoCambio.Locked = False
End If

vPaso = False

Call sbConsulta(gIVR_Transito.TituloId)
Call sbInicializa

If pConsultaExternaId > 0 Then
    Call sbConsuta_Transac(pConsultaExternaId)
End If

Call sbAdquisicion_Load

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbConsuta_Transac(pTransacId As Long)

strSQL = "exec spIVR_FONDO_MOV_CONSULTA " & txtInversionId.Text & "," & pTransacId
Call OpenRecordSet(rs, strSQL)

pConsultaExternaId = 0

txtTransacId.Text = rs!FI_Transac
txtSeqId.Text = rs!Seq_Id

Select Case rs!Estado
    Case "S"
        txtEstado.Text = "Solicitud"
        txtEstado.Tag = "S"
    Case Else
        txtEstado.Text = "Registrado"
        txtEstado.Tag = "R"
End Select

Select Case rs!Tipo_Mov
    Case "A"
        gIVR_Transito.Concepto = "FI_APO"
        gIVR_Transito.Tipo = "F"
        gIVR_Transito.Monto = 0
        gIVR_Transito.TipoMov = "C"
            
    Case "R"
        gIVR_Transito.Concepto = "FI_RET"
        gIVR_Transito.Tipo = "F"
        gIVR_Transito.Monto = 0
        gIVR_Transito.TipoMov = "D"
    Case Else
        gIVR_Transito.Concepto = "FI_APO"
        gIVR_Transito.Tipo = "F"
        gIVR_Transito.Monto = 0
        gIVR_Transito.TipoMov = "C"
End Select

If gIVR_Transito.Concepto = "FI_APO" Then
     scGestion.Caption = "Gestión:   Aportación al Fondo"
Else
     scGestion.Caption = "Gestión:   Retiro del Fondo"
End If


txtDocumento.Text = rs!Documento & ""

dtpFi_FechaMovimiento.Value = rs!Fecha
txtFi_Movimiento.Text = Format(rs!Mnt_Principal, "Standard")


txtFi_ParticipacionNo.Text = rs!N_Participaciones
txtFi_ParticipacionValor.Text = rs!V_Participaciones

txtFi_Movimiento.SetFocus

txtNotas.Text = rs!NOTAS

Call sbFondos_Cal_Refresh

lswAd.ListItems.Clear
txtAd_Pendiente.Text = Format(0, "Standard")
txtAd_Registrado.Text = Format(0, "Standard")


rs.Close

End Sub

Public Sub sbConsultaExterna(pTransacId As Long)

pConsultaExternaId = pTransacId

End Sub


Private Sub sbConsulta(pTituloId As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vIVR_INVERSIONES" _
       & " Where Titulo_ID = " & pTituloId
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    txtInversionId.Text = rs!TITULO_ID
    
    dtpFi_FechaMovimiento.Value = DateAdd("d", 1, vFecha)
    
    txtAdministrador.Text = rs!Administrador_Desc
    txtAdministrador.Tag = rs!Cod_Administrador
    
    txtInstrumento.Text = rs!Instrumento_Desc
    txtInstrumento.Tag = rs!Cod_Instrumento
    
    txtPortafolio.Text = rs!Portafolio_Desc
    txtPortafolio.Tag = rs!Cod_Portafolio
    
        
    txtFi_Movimiento.Text = 0

    txtFi_ValorActual.Text = Format(rs!Monto_Principal, "Standard")
    txtFi_ParticipacionNo.Text = rs!Participacion_Numero
    txtFi_ParticipacionValor.Text = rs!Participacion_Valor
    
    If vDivisaLocaL = IIf(IsNull(rs!Cod_Divisa), gIVR_Transito.Divisa, rs!Cod_Divisa) Then
        txtTipoCambio.Text = "1"
        txtTipoCambio.Locked = True
    Else
        txtTipoCambio.Locked = False
    End If
    
Else
  Me.MousePointer = vbDefault
  MsgBox "No se Localizó el registro!", vbExclamation
End If
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbInicializa()

vPaso = True


txtTransacId.Text = "0"
txtSeqId.Text = "0"
txtEstado.Text = "Solicitud"

txtDocumento.Text = ""

dtpFi_FechaMovimiento.Value = DateAdd("d", 1, vFecha)
txtFi_Movimiento.Text = Format(0, "Standard")


txtFi_ParticipacionNo.Text = 1
txtFi_ParticipacionValor.Text = 1

txtFi_Movimiento.SetFocus

txtNotas.Text = ""

lswAd.ListItems.Clear
txtAd_Pendiente.Text = Format(0, "Standard")
txtAd_Registrado.Text = Format(0, "Standard")


lswAd.ListItems.Clear

vPaso = False

dtpFi_FechaMovimiento.SetFocus

End Sub


Private Sub sbGuardar(Optional pEstado As String = "S")
On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pTipoMov As String, pMonto As Currency, pDivisa As String, pTipoCambio As Currency
Dim pCuenta As String, pBancoId As Long


With gIVR_Transito


If gIVR_Transito.Concepto = "FI_APO" Then
    pMonto = CCur(txtFi_Movimiento.Text)
    pTipoMov = "A"
Else
    pMonto = CCur(txtFi_Movimiento.Text) * -1
    pTipoMov = "R"
End If

strSQL = "exec spIVR_FONDO_MOV_REGISTRA " & txtInversionId.Text & ", '" & pTipoMov _
       & "','" & .Concepto _
       & "','" & Format(dtpFi_FechaMovimiento.Value, "yyyy/mm/dd") & "','" & txtDocumento.Text _
       & "', " & pMonto & ", 0" _
       & " , " & CDbl(txtFi_ParticipacionNo.Text) _
       & " , " & CDbl(txtFi_ParticipacionValor.Text) _
       & " ,'" & txtDivisa.Text & "' ," & CCur(txtTipoCambio.Text) _
       & " ,'" & txtNotas.Text _
       & "','" & glogon.Usuario & "', '" & pEstado & "', " & txtTransacId.Text & ", " & txtSeqId.Text
End With

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
    txtTransacId.Text = rs!Transac_Id
    txtSeqId.Text = rs!Seq_Id
    
    If rs!Estado = "S" Then
        txtEstado.Text = "Solicitud"
    Else
        txtEstado.Text = "Procesada"
    End If
Else
     Call sbInicializa
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub sbFondos_Cal_Refresh()

On Error GoTo vError

If Not IsNumeric(txtFi_Movimiento.Text) Then
    txtFi_Movimiento.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionNo.Text) Then
    txtFi_ParticipacionNo.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionValor.Text) Then
    txtFi_ParticipacionValor.Text = "1"
End If

If Not IsNumeric(txtTipoCambio.Text) Then
    txtTipoCambio.Text = "1"
End If

    
If Not IsNumeric(txtAd_Registrado.Text) Then
    txtAd_Registrado.Text = "0"
End If

If Not IsNumeric(txtAd_Pendiente.Text) Then
    txtAd_Pendiente.Text = "0"
End If
    
'Formato
txtFi_ParticipacionNo.Text = Format(CCur(txtFi_Movimiento.Text) / CDbl(txtFi_ParticipacionValor), "###,###,###,##0.0000000000")
txtFi_ParticipacionValor.Text = Format(CDbl(txtFi_ParticipacionValor), "Standard")
txtFi_Movimiento.Text = Format(CCur(txtFi_Movimiento.Text), "Standard")


txtTipoCambio.Text = Format(CCur(txtTipoCambio.Text), "###,##0.000000")
txtImporteLocal.Text = Format(CCur(txtFi_Movimiento.Text) * fxSys_Tipo_Cambio_Apl(CCur(txtTipoCambio.Text)), "Standard")

txtAd_Pendiente.Text = Format(CCur(txtImporteLocal.Text) - CCur(txtAd_Registrado.Text), "Standard")

Exit Sub

vError:

End Sub


Private Sub txtFi_Movimiento_GotFocus()
On Error GoTo vError
    txtFi_Movimiento.Text = Format(CCur(txtFi_Movimiento.Text), "Standard")
vError:
End Sub

Private Sub txtFi_Movimiento_LostFocus()
Call sbFondos_Cal_Refresh
End Sub


Private Sub txtFi_ParticipacionValor_LostFocus()
Call sbFondos_Cal_Refresh
End Sub


Private Sub txtTipoCambio_GotFocus()
On Error GoTo vError
    txtTipoCambio.Text = Format(CCur(txtTipoCambio.Text), "###,##0.000000")
vError:
End Sub

Private Sub txtTipoCambio_LostFocus()
Call sbFondos_Cal_Refresh
End Sub
