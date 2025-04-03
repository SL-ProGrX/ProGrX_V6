VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRH_Accion_Personal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Registro de Acción de Personal"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox gbAplica 
      Height          =   2415
      Left            =   240
      TabIndex        =   37
      Top             =   6240
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   4260
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   615
         Left            =   7320
         TabIndex        =   38
         Top             =   1680
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmRH_Accion_Personal.frx":0000
      End
      Begin XtremeSuiteControls.ComboBox cboTipoAccion 
         Height          =   312
         Left            =   1560
         TabIndex        =   39
         Top             =   240
         Width           =   6132
         _Version        =   1441793
         _ExtentX        =   10821
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   912
         Left            =   1560
         TabIndex        =   41
         Top             =   600
         Width           =   6132
         _Version        =   1441793
         _ExtentX        =   10816
         _ExtentY        =   1609
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
      Begin XtremeSuiteControls.DateTimePicker dtpAccion 
         Height          =   312
         Left            =   3480
         TabIndex        =   44
         Top             =   1680
         Width           =   1332
         _Version        =   1441793
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
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Real de la Acción de Personal:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   1560
         TabIndex        =   43
         Top             =   1680
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
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
         Height          =   372
         Index           =   4
         Left            =   360
         TabIndex        =   42
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8640
      Top             =   120
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1812
      _Version        =   1441793
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEmpleadoId 
      Height          =   312
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1812
      _Version        =   1441793
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   5052
      _Version        =   1441793
      _ExtentX        =   8911
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbAccionPersonal 
      Height          =   5055
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   8916
      _StockProps     =   79
      ForeColor       =   4210752
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
      Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
         Height          =   252
         Index           =   0
         Left            =   7800
         TabIndex        =   7
         Top             =   360
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
         Height          =   252
         Index           =   1
         Left            =   7800
         TabIndex        =   8
         Top             =   720
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
         Height          =   252
         Index           =   2
         Left            =   7800
         TabIndex        =   9
         Top             =   1080
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
         Height          =   252
         Index           =   3
         Left            =   7800
         TabIndex        =   10
         Top             =   1560
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroCod 
         Height          =   312
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
         Height          =   312
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9758
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
         Height          =   312
         Left            =   1440
         TabIndex        =   13
         Top             =   720
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
         Height          =   312
         Left            =   2160
         TabIndex        =   14
         Top             =   720
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9758
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
         Height          =   312
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtSecDesc 
         Height          =   312
         Left            =   2160
         TabIndex        =   16
         Top             =   1080
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9758
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPuestoCod 
         Height          =   312
         Left            =   1440
         TabIndex        =   17
         Top             =   1560
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtPuestoDesc 
         Height          =   312
         Left            =   2160
         TabIndex        =   18
         Top             =   1560
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9758
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboNomina 
         Height          =   312
         Left            =   1440
         TabIndex        =   19
         Top             =   2880
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      End
      Begin XtremeSuiteControls.ComboBox cboContrato 
         Height          =   312
         Left            =   1440
         TabIndex        =   20
         Top             =   3240
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      End
      Begin XtremeSuiteControls.ComboBox cboJornada 
         Height          =   312
         Left            =   1440
         TabIndex        =   21
         Top             =   4200
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      End
      Begin XtremeSuiteControls.ComboBox cboVacaciones 
         Height          =   312
         Left            =   1440
         TabIndex        =   22
         Top             =   4560
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      End
      Begin XtremeSuiteControls.FlatEdit txtSalario 
         Height          =   312
         Left            =   5280
         TabIndex        =   23
         Top             =   2400
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
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
      Begin XtremeSuiteControls.DateTimePicker dtpContrato 
         Height          =   312
         Left            =   6360
         TabIndex        =   24
         Top             =   3600
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   312
         Left            =   5280
         TabIndex        =   25
         Top             =   2040
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4260
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
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Salario:"
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
         Left            =   4320
         TabIndex        =   36
         Top             =   2400
         Width           =   1932
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Régimen de Vacaciones"
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
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Top             =   4560
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Jornada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   4200
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nómina"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   2880
         Width           =   1092
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Centro"
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
         Index           =   12
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   972
      End
      Begin VB.Label lblDepartamento 
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
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label lblSeccion 
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
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   1572
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Puesto"
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
         TabIndex        =   28
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label lblVencimiento 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
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
         Left            =   4320
         TabIndex        =   27
         Top             =   3600
         Width           =   1452
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa:"
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
         Left            =   4320
         TabIndex        =   26
         Top             =   2040
         Width           =   1932
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   3960
      TabIndex        =   5
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Id. Empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmRH_Accion_Personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vScroll As Boolean

Private Sub btnAplicar_Click()
If vPaso Then Exit Sub

If cboTipoAccion.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset
Dim Boleta As String

'Validacion


On Error GoTo vError


If txtEmpleadoId.Text = "" Then
    MsgBox "No se ha indicado a ningún Empleado!", vbExclamation
    Exit Sub
End If

If CCur(txtSalario.Text) < CCur(txtSalario.Tag) Then
    MsgBox "Salario no puede ser inferior al anterior!", vbExclamation
    Exit Sub
End If


strSQL = "exec spRH_Accion_Personal_Registro '" & txtEmpleadoId.Text & "','" & cboTipoAccion.ItemData(cboTipoAccion.ListIndex) _
        & "','" & txtNotas.Text & "','" & glogon.Usuario & "'" _
        & ",'A','" & txtPuestoCod.Text & "','" & txtCentroCod.Text & "','" & txtDeptCodigo.Text & "','" & txtSecCodigo.Text _
        & "'," & CCur(txtSalario.Text) & ",'" & cboNomina.ItemData(cboNomina.ListIndex) _
        & "','" & Format(dtpAccion.Value, "yyyy/mm/dd") & "', '" & cboContrato.ItemData(cboContrato.ListIndex) _
        & "','" & cboJornada.ItemData(cboJornada.ListIndex) & "','" & cboVacaciones.ItemData(cboVacaciones.ListIndex) & "'"
        
Call OpenRecordSet(rs, strSQL)
    Boleta = rs!Accion_Personal
rs.Close

'Print Boleta
Call sbBoleta_Accion_Personal(Boleta)

Me.MousePointer = vbDefault

MsgBox "Accion de Personal, registrada satisfactoriamente!", vbInformation

Call sbLimpia


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub FlatScroll_Laboral_Change(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCodigo As String, vColumna As String, vChar As String, vFiltroAdd As String
Dim txtCodigo As Object, txtDesc As Object

On Error GoTo vError

If Not vScroll Then Exit Sub

vChar = "'"
vFiltroAdd = ""

Select Case Index
   Case 0 'Centro
        vCodigo = txtCentroCod.Text
        vColumna = "COD_CENTRO"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_CENTRO_TRABAJO"
        
        Set txtCodigo = txtCentroCod
        Set txtDesc = txtCentroDesc
    
    Case 1 'Departamentos
        vCodigo = txtDeptCodigo.Text
        vColumna = "COD_DEPARTAMENTO"
        vFiltroAdd = " AND COD_CENTRO = '" & txtCentroCod.Text & "'"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_Departamentos"
        
        Set txtCodigo = txtDeptCodigo
        Set txtDesc = txtDeptDesc
        
        
    Case 2 'Secciones
        vCodigo = txtSecCodigo.Text
        vColumna = "COD_SECCION"
        vFiltroAdd = " AND COD_CENTRO = '" & txtCentroCod.Text & "' AND COD_DEPARTAMENTO = '" & txtDeptCodigo.Text & "'"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_SECCIONES"
        
        Set txtCodigo = txtSecCodigo
        Set txtDesc = txtSecDesc

    
    Case 3 'Puesto
        vCodigo = txtPuestoCod.Text
        
        vColumna = "COD_PUESTO"
        vFiltroAdd = " AND ACTIVO = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_PUESTOS"
        
        Set txtCodigo = txtPuestoCod
        Set txtDesc = txtPuestoDesc
    
    
End Select

If vScroll Then
    
    If FlatScroll_Laboral(Index).Value = 1 Then
       strSQL = strSQL & " where " & vColumna & " > " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " asc"
    Else
       strSQL = strSQL & " where " & vColumna & " < " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Codigo
      txtDesc.Text = rs!Descripcion

    End If
    rs.Close
End If



vScroll = False
FlatScroll_Laboral(Index).Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()

vModulo = 23

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


vScroll = True
 
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpia()

txtEmpleadoId.Text = ""
txtIdentificacion.Text = ""
txtNombre.Text = ""


txtCentroCod.Text = ""
txtCentroDesc.Text = ""
txtDeptCodigo.Text = ""
txtDeptDesc.Text = ""
txtSecCodigo.Text = ""
txtSecDesc.Text = ""

txtPuestoCod.Text = ""
txtPuestoDesc.Text = ""

txtNotas.Text = ""
txtSalario.Text = Format(0, "Standard")



End Sub


Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vPaso = True

'Tipo de Accion
    strSQL = "select Tipo_Accion as Idx, rtrim(Descripcion) as ItmX from RH_ACCION_PERSONAL_TIPOS"
    Call sbCbo_Llena_New(cboTipoAccion, strSQL, False, True)

'Nomina
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)

'Divisa
    strSQL = "select COD_DIVISA as Idx, rtrim(Descripcion) as ItmX from vSys_Divisas"
    Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

'Jornada
    strSQL = "select JORNADA_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_JORNADAS_TIPOS"
    Call sbCbo_Llena_New(cboJornada, strSQL, False, True)

'Contratos
    strSQL = "Select CONTRATO_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_CONTRATOS_TIPOS"
    Call sbCbo_Llena_New(cboContrato, strSQL, False, True)

'Vacaciones
    strSQL = "Select COD_VACA_REGIMEN as Idx, rtrim(Descripcion) as ItmX from RH_VACACIONES_REGIMEN"
    Call sbCbo_Llena_New(cboVacaciones, strSQL, False, True)


dtpAccion.Value = fxFechaServidor

vPaso = False

txtEmpleadoId.SetFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub sbBusca()
   gBusquedas.Convertir = "N"
   gBusquedas.Col1Name = "Empleado Id"
   gBusquedas.Col2Name = "Persona Id"
   gBusquedas.Col3Name = "Nombre"
   gBusquedas.Columna = "Empleado_ID"
   gBusquedas.Orden = "Empleado_ID"
   gBusquedas.Consulta = "Select Empleado_ID, Identificacion, Nombre_Completo From Rh_Personas"
   
   gBusquedas.Filtro = " and ESTADO_PERSONA = 'A'"
   
   frmBusquedas.Show vbModal
   
   txtEmpleadoId.Text = gBusquedas.Resultado
   txtIdentificacion.Text = Trim(gBusquedas.Resultado2)
   txtNombre.Text = gBusquedas.Resultado3
    
   Call sbConsulta
    
End Sub

Public Sub sbConsulta_Externa(pEmpleadoId As String)

txtEmpleadoId.Text = pEmpleadoId
Call sbConsulta

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vRH_Personas" _
       & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtEmpleadoId.Text = rs!Empleado_ID
    txtIdentificacion.Text = rs!IDENTIFICACION
    txtNombre.Text = rs!NOMBRE_COMPLETO
    
    txtCentroCod.Text = rs!Cod_Centro
    txtCentroDesc.Text = rs!CentroDesc
    txtDeptCodigo.Text = rs!Cod_Departamento
    txtDeptDesc.Text = rs!DepartamentoDesc
    txtSecCodigo.Text = rs!Cod_Seccion
    txtSecDesc.Text = rs!SeccionDesc
    
    txtPuestoCod.Text = rs!Cod_Puesto
    txtPuestoDesc.Text = rs!PuestoDesc
    
    txtSalario.Text = Format(rs!SALARIO_ORDINARIO, "Standard")
    txtSalario.Tag = Format(rs!SALARIO_ORDINARIO, "Standard")
    
   Call sbCboAsignaDato(cboDivisa, rs!DivisaDesc, True, rs!cod_Divisa)
   
   Call sbCboAsignaDato(cboNomina, rs!NominaDesc, True, rs!COD_NOMINA)
   Call sbCboAsignaDato(cboContrato, rs!ContratoDesc, True, rs!Contrato_Tipo)
   Call sbCboAsignaDato(cboJornada, rs!JornadaDesc, True, rs!Jornada_Tipo)
   Call sbCboAsignaDato(cboVacaciones, rs!VacacionesDesc, True, rs!Cod_Vaca_Regimen)
   
   If Not IsNull(rs!Contrato_Vencimiento) Then
        dtpContrato.Value = rs!Contrato_Vencimiento
        dtpContrato.Visible = True
   Else
        dtpContrato.Visible = False
   End If
   lblVencimiento.Visible = dtpContrato.Visible
   
    
Else
    'Todo
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtCentroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion,desc_Corta from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtCentroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion,desc_Corta from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus

If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "cod_departamento"
    gBusquedas.Orden = "cod_departamento"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"
  
   
  
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub


Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"

  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtEmpleadoId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusca
End Sub




Private Sub txtPuestoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPuestoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_PUESTO"
  gBusquedas.Orden = "COD_PUESTO"
  gBusquedas.Consulta = "select COD_PUESTO,descripcion from Rh_Puestos"
  gBusquedas.Filtro = ""
        
  frmBusquedas.Show vbModal
  txtPuestoCod.Text = gBusquedas.Resultado
  txtPuestoDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtPuestoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDivisa.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_PUESTO,descripcion from Rh_Puestos"
  gBusquedas.Filtro = ""
        
  frmBusquedas.Show vbModal
  txtPuestoCod.Text = gBusquedas.Resultado
  txtPuestoDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtSalario_GotFocus()
On Error GoTo vError

txtSalario.Text = CCur(txtSalario.Text)

vError:
End Sub

Private Sub txtSalario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then cboNomina.SetFocus
End Sub

Private Sub txtSalario_LostFocus()
On Error GoTo vError

txtSalario.Text = Format(CCur(txtSalario.Text), "Standard")

vError:

End Sub


Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then

        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from Rh_Secciones"
        gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text _
                  & "' and cod_departamento = '" & txtDeptCodigo & "'"
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If
End Sub

