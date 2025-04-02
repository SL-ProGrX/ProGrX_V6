VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Constancias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Constancias de Deudas"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9720
      Top             =   480
   End
   Begin XtremeSuiteControls.GroupBox gbUniversidad 
      Height          =   3615
      Left            =   5640
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   6376
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cboUniversidad 
         Height          =   345
         Left            =   0
         TabIndex        =   24
         Top             =   360
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.ComboBox cboNivel 
         Height          =   345
         Left            =   0
         TabIndex        =   25
         Top             =   960
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.ComboBox cboCarrera 
         Height          =   345
         Left            =   0
         TabIndex        =   30
         Top             =   1560
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.ComboBox cboEspecialidad 
         Height          =   345
         Left            =   0
         TabIndex        =   31
         Top             =   2160
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
         Height          =   345
         Left            =   0
         TabIndex        =   32
         Top             =   3120
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   609
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboParentesco 
         Height          =   345
         Left            =   2040
         TabIndex        =   36
         Top             =   2760
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiarioId 
         Height          =   345
         Left            =   0
         TabIndex        =   34
         Top             =   2760
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   609
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCiclo 
         Height          =   345
         Left            =   2040
         TabIndex        =   37
         Top             =   960
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.FlatEdit txtCicloAnio 
         Height          =   345
         Left            =   3960
         TabIndex        =   38
         Top             =   960
         Width           =   630
         _Version        =   1441793
         _ExtentX        =   1111
         _ExtentY        =   609
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
         Text            =   "2024"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   39
         Top             =   720
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ciclo:"
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
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   35
         Top             =   2520
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Parentesco:"
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
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   33
         Top             =   2520
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Beneficiario:"
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
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Universidad:"
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
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   28
         Top             =   720
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Grado Académico:"
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
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Carrera:"
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
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   26
         Top             =   1920
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Especialidad:"
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
   End
   Begin XtremeSuiteControls.CheckBox chkIdAlterna 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   7800
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Utiliza Identificación alterna?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtDirigidoA 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   5880
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "A quién interese"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   8280
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Informe"
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
      Appearance      =   16
      Picture         =   "frmCR_Constancias.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   615
      Left            =   8760
      TabIndex        =   6
      Top             =   8280
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   1080
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
      Appearance      =   16
      Picture         =   "frmCR_Constancias.frx":07BC
   End
   Begin XtremeSuiteControls.FlatEdit txtEmitidoPor 
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   6480
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Responsable"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPuesto 
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   7080
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Puesto"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   8280
      TabIndex        =   12
      Top             =   7800
      Width           =   1335
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
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   14
      Top             =   1920
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Constancia de Deudas"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   15
      Top             =   2280
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Autorización DTR"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   16
      ToolTipText     =   "Centro de Información Crediticia de SUGEF"
      Top             =   2640
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Autorización CIC"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   17
      Top             =   3000
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Registro de Firmas"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   18
      Top             =   3360
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Declación Jurada de Domicilio"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   19
      Top             =   3720
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Autorización de Deducción para Pensionados"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   6
      Left            =   1200
      TabIndex        =   20
      Top             =   4080
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Certificación de Cuenta IBAN"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   21
      Top             =   1920
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Carta para Universidades"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboCtaIBAN 
      Height          =   345
      Left            =   1200
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   13
      Top             =   7800
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Corte Intereses:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   5880
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Dirigido a :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   6480
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Emitido por :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   7080
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Puesto :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "CEDULA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   7935
      _Version        =   1441793
      _ExtentX        =   13996
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "NOMBRE_COMPLETO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Constancias de Créditos"
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
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmCR_Constancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mCedula As String, vPaso As Boolean

Private Sub btnCerrar_Click()
 Unload Me
End Sub



Private Sub btnReporte_Click()
Dim pGestion As String, pNotas As String

On Error GoTo vError

Me.MousePointer = vbHourglass

pGestion = "99"
pNotas = ""

With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = False
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Sistema de Crédito"

   .Connect = glogon.ConectRPT
   

   Select Case True
       Case OptX.Item(0).Value 'Constancia de Deudas
       
              pGestion = "01"
       
             strSQL = "{SOCIOS.CEDULA} = '" & mCedula & "'"
             .ReportFileName = SIFGlobal.fxPathReportes("Sys_EstadoConstancia.rpt")
              
             .Formulas(0) = "fxDirigido='" & txtDirigidoA.Text & "'"
             .Formulas(1) = "fxEmite='" & txtEmitidoPor.Text & "'"
             .Formulas(2) = "fxPuesto='" & txtPuesto.Text & "'"
             
             
             .SelectionFormula = strSQL
            
             .SubreportToChange = "sbTexto"
             .StoredProcParam(0) = mCedula
             .StoredProcParam(1) = chkIdAlterna.Value
             .StoredProcParam(2) = 4
              
             .SubreportToChange = "sbCreditos"
             .Formulas(0) = "fxCorte = 'Fecha de Corte para Cálculo de Intereses : " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
             .StoredProcParam(0) = mCedula
             .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy-mm-dd")
        
   
       Case OptX.Item(1).Value 'Autorización DTR
              pGestion = "02"
             
             .ReportFileName = SIFGlobal.fxPathReportes("Sys_Autorizacion_DTR.rpt")
             .StoredProcParam(0) = mCedula
             
      
       Case OptX.Item(2).Value 'Autorización SIC
              pGestion = "03"
       
             .ReportFileName = SIFGlobal.fxPathReportes("Sys_Autorizacion_SIC.rpt")
             .StoredProcParam(0) = mCedula

             
       Case OptX.Item(3).Value 'Registro de Firmas
              pGestion = "04"
             
             .ReportFileName = SIFGlobal.fxPathReportes("Sys_Registro_Firmas.rpt")
            
             .StoredProcParam(0) = mCedula
       
       Case OptX.Item(4).Value 'Declación Jurada de Domicilio
              pGestion = "05"
             
             .ReportFileName = SIFGlobal.fxPathReportes("Sys_Declaracion_Jurada_Domicilio.rpt")
            
             .Formulas(0) = "fxEmite='" & txtEmitidoPor.Text & "'"
             .StoredProcParam(0) = mCedula
            
            

       Case OptX.Item(5).Value 'Autorizacion de Deduccion para Pensionados
              pGestion = "07"
              
              .ReportFileName = SIFGlobal.fxPathReportes("Sys_Autorizacion_Deduccion_Pensionados.rpt")
            
             .StoredProcParam(0) = mCedula
            
       
       Case OptX.Item(6).Value 'Certificacion de Cuenta IBAN
              pGestion = "08"
              pNotas = cboCtaIBAN.Text
              
             .ReportFileName = SIFGlobal.fxPathReportes("Sys_Certificacion_IBAN.rpt")
            
             .Formulas(0) = "fxEmite='" & txtEmitidoPor.Text & "'"
             .Formulas(1) = "fxIBAN='" & cboCtaIBAN.ItemData(cboCtaIBAN.ListIndex) & "'"
             
             .StoredProcParam(0) = mCedula
             
             .SubreportToChange = "sbLista"
             .StoredProcParam(0) = mCedula
             .StoredProcParam(1) = cboCtaIBAN.ItemData(cboCtaIBAN.ListIndex)
        
       
       Case OptX.Item(7).Value 'Carta a Universidades
             pGestion = "09"
             pNotas = cboUniversidad.Text & ", " & cboNivel.Text & ", " & cboCarrera.Text & ", " & cboCiclo.Text & " " & txtCicloAnio.Text
             
             .ReportFileName = SIFGlobal.fxPathReportes("Sys_Carta_Convenio_Universidad.rpt")
            
             .Formulas(0) = "fxEmite='" & txtEmitidoPor.Text & "'"
             
             If txtBeneficiarioId.Text = "" Then
                txtBeneficiarioId.Text = "NA"
                txtBeneficiario.Text = "NA"
             End If
             
             .StoredProcParam(0) = mCedula
             .StoredProcParam(1) = cboUniversidad.ItemData(cboUniversidad.ListIndex)
             .StoredProcParam(2) = cboNivel.ItemData(cboNivel.ListIndex)
             .StoredProcParam(3) = cboCarrera.ItemData(cboCarrera.ListIndex)
             .StoredProcParam(4) = cboEspecialidad.ItemData(cboEspecialidad.ListIndex)
             .StoredProcParam(5) = txtBeneficiarioId.Text
             .StoredProcParam(6) = txtBeneficiario.Text
             .StoredProcParam(7) = cboParentesco.ItemData(cboParentesco.ListIndex)
             .StoredProcParam(8) = cboCiclo.Text
             .StoredProcParam(9) = txtCicloAnio.Text
             .StoredProcParam(10) = txtEmitidoPor.Text
             .StoredProcParam(11) = glogon.Usuario
             
             If txtBeneficiarioId.Text = "NA" Then
                txtBeneficiarioId.Text = ""
                txtBeneficiario.Text = ""
             End If
             
             
   End Select
   
   .Action = 1
End With

'Bitacora Especial
strSQL = "exec spSys_Bitacora_Operaciones_Registra '" & pGestion & "','" & mCedula & "','" & pNotas & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboCarrera_Click()
If vPaso Then Exit Sub

vPaso = True

strSQL = "exec spSys_Educacion_List 'E', '" & cboCarrera.ItemData(cboCarrera.ListIndex) & "'"
Call sbCbo_Llena_New(cboEspecialidad, strSQL, False, True)

cboEspecialidad.AddItem "No Indica"
cboEspecialidad.ItemData(cboEspecialidad.ListCount - 1) = "NA"
cboEspecialidad.Text = "No Indica"

vPaso = False

End Sub

Private Sub cboUniversidad_Click()
If vPaso Then Exit Sub

vPaso = True

'strSQL = "exec spSys_Educacion_List 'N', '" & cboUniversidad.ItemData(cboUniversidad.ListIndex) & "'"
strSQL = "exec spSys_Educacion_List 'N', ''"
Call sbCbo_Llena_New(cboNivel, strSQL, False, True)

'strSQL = "exec spSys_Educacion_List 'C', '" & cboUniversidad.ItemData(cboUniversidad.ListIndex) & "'"
strSQL = "exec spSys_Educacion_List 'C', ''"
Call sbCbo_Llena_New(cboCarrera, strSQL, False, True)

vPaso = False

Call cboCarrera_Click


End Sub

Private Sub Form_Load()

vModulo = 3

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

scMain.Item(0).Caption = GLOBALES.gTag
scMain.Item(1).Caption = GLOBALES.gTag2

mCedula = GLOBALES.gTag

dtpCorte.Value = GLOBALES.gTag3

txtPuesto.Text = ""


Exit Sub

vError:

End Sub



Private Sub sbInicializa()

On Error GoTo vError


cboCiclo.Clear
cboCiclo.AddItem "I   Quatrimestre"
cboCiclo.AddItem "II  Quatrimestre"
cboCiclo.AddItem "III Quatrimestre"
cboCiclo.AddItem "IV  Quatrimestre"

cboCiclo.AddItem "I   Semestre"
cboCiclo.AddItem "II  Semestre"

cboCiclo.Text = "I   Quatrimestre"

txtCicloAnio.Text = Year(fxFechaServidor)

strSQL = "select descripcion from Usuarios where Nombre  = '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
  txtEmitidoPor.Text = rs!Descripcion
rs.Close

'IBAN Internas
cboCtaIBAN.Clear

strSQL = "exec spSys_Cuenta_SINPE '" & GLOBALES.gTag & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 

 cboCtaIBAN.AddItem rs!IBAN_MASK & ""
 cboCtaIBAN.ItemData(cboCtaIBAN.ListCount - 1) = CStr(rs!IBAN)
 
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboCtaIBAN.Text = rs!IBAN_MASK & ""
End If
rs.Close

cboCtaIBAN.AddItem "TODAS"
cboCtaIBAN.ItemData(cboCtaIBAN.ListCount - 1) = "TODAS"


 strSQL = "select rtrim(cod_Parentesco) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from sys_Parentescos where activo = 1"
 Call sbCbo_Llena_New(cboParentesco, strSQL, False, True)
 

Call OptX_Click(0)


Exit Sub

vError:


End Sub

Public Sub sbCrdConstancia(pCedula As String, pCorte As Date, Optional pPrinter As Integer = 0)
Dim rs As New ADODB.Recordset, strSQL As String

Screen.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
     .Reset
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Cuentas Corrientes"
     
     .Connect = glogon.ConectRPT
     
     strSQL = "select CONSTANCIA_CRD_ENCABEZADO from sif_Empresa"
     Call OpenRecordSet(rs, strSQL)
        .Formulas(0) = "formula4 = '" & Trim(rs!CONSTANCIA_CRD_ENCABEZADO & "") & "'"
     rs.Close
          
     
     strSQL = "select descripcion from Usuarios where Nombre  = '" & glogon.Usuario & "'"
     Call OpenRecordSet(rs, strSQL)
        .Formulas(1) = "fxUsuario = '" & Trim(rs!Descripcion & "") & "'"
     rs.Close
               
     
     .ReportFileName = SIFGlobal.fxPathReportes("Sys_EstadoConstancia.rpt")
     .SelectionFormula = "{SOCIOS.CEDULA} = '" & pCedula & "'"
     
       .SubreportToChange = "sbCreditos"
       .Formulas(0) = "fxCorte = 'Fecha de Corte para Cálculo de Intereses : " & Format(pCorte, "dd/mm/yyyy") & "'"
       .StoredProcParam(0) = pCedula
       .StoredProcParam(1) = Format(pCorte, "yyyy-mm-dd")
    
     If pPrinter = 1 Then .Destination = crptToPrinter
     .PrintReport
End With

Screen.MousePointer = vbDefault
Exit Sub

vError:
 Screen.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbUniversidades_Load()

vPaso = True
strSQL = "exec spSys_Educacion_List 'U', Null"
Call sbCbo_Llena_New(cboUniversidad, strSQL, False, True)
vPaso = False
End Sub

Private Sub OptX_Click(Index As Integer)

 lblTitulo.Caption = OptX.Item(Index).Caption

cboCtaIBAN.Visible = False
gbUniversidad.Visible = False

If Index = 6 Then
    cboCtaIBAN.Visible = True
End If

If Index = 7 Then
    gbUniversidad.Visible = True
    Call sbUniversidades_Load
End If




End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub


Private Sub txtBeneficiarioId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Nombre"
    gBusquedas.Consulta = "Select Identificacion, Nombre from vSys_Padron_Nacional"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtBeneficiarioId.Text = gBusquedas.Resultado
        txtBeneficiario.Text = gBusquedas.Resultado2
    End If
End If
End Sub

Private Sub txtBeneficiarioId_LostFocus()
txtBeneficiario.Text = fxPadron_Nacional_Nombre(txtBeneficiarioId.Text)
End Sub
