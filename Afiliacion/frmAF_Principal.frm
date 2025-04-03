VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.TaskPanel.v24.0.0.ocx"
Begin VB.Form frmAF_Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Personas"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   Icon            =   "frmAF_Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   12015
   Begin XtremeTaskPanel.TaskPanel tpMain 
      Height          =   7365
      Left            =   0
      TabIndex        =   72
      Top             =   1680
      Width           =   2760
      _Version        =   1572864
      _ExtentX        =   4868
      _ExtentY        =   12991
      _StockProps     =   64
      VisualTheme     =   17
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeSuiteControls.GroupBox gbDimeX 
      Height          =   4095
      Left            =   11880
      TabIndex        =   170
      Top             =   4680
      Visible         =   0   'False
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
      _ExtentY        =   7223
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox gbDimex_Log 
         Height          =   975
         Index           =   0
         Left            =   960
         TabIndex        =   177
         Top             =   1800
         Width           =   6735
         _Version        =   1572864
         _ExtentX        =   11880
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Registrado: "
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.FlatEdit txtDimex_RUsuario 
            Height          =   330
            Left            =   1320
            TabIndex        =   180
            Top             =   600
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDimex_RFecha 
            Height          =   330
            Left            =   3600
            TabIndex        =   181
            Top             =   600
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label lblDimSec 
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   179
            Top             =   360
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblDimSec 
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   178
            Top             =   360
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            Alignment       =   2
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox chkDimex_Activo 
         Height          =   255
         Left            =   5760
         TabIndex        =   176
         Top             =   840
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activo ?"
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
      End
      Begin XtremeSuiteControls.PushButton btnDimex 
         Height          =   330
         Index           =   0
         Left            =   3960
         TabIndex        =   173
         Top             =   1320
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Actualiza"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Principal.frx":6852
      End
      Begin XtremeSuiteControls.PushButton btnDimex 
         Height          =   210
         Index           =   1
         Left            =   8760
         TabIndex        =   174
         ToolTipText     =   "Cierra: Actualizacion de Dimex"
         Top             =   480
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   450
         _ExtentY        =   370
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAF_Principal.frx":6F79
      End
      Begin XtremeSuiteControls.FlatEdit txtDimex_Nuevo 
         Height          =   330
         Left            =   3360
         TabIndex        =   175
         Top             =   840
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbDimex_Log 
         Height          =   975
         Index           =   1
         Left            =   960
         TabIndex        =   182
         Top             =   2880
         Width           =   6735
         _Version        =   1572864
         _ExtentX        =   11880
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Actualizado: "
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.FlatEdit txtDimex_AUsuario 
            Height          =   330
            Left            =   1320
            TabIndex        =   183
            Top             =   600
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDimex_AFecha 
            Height          =   330
            Left            =   3600
            TabIndex        =   184
            Top             =   600
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label lblDimSec 
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   186
            Top             =   360
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            Alignment       =   2
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblDimSec 
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   185
            Top             =   360
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha"
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label lblDimSec 
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   172
         Top             =   600
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "DIMEX"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   0
         TabIndex        =   171
         Top             =   0
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Actualización del DIMEX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   10080
      Top             =   720
   End
   Begin XtremeSuiteControls.PushButton btnIngresa 
      Height          =   330
      Index           =   0
      Left            =   6960
      TabIndex        =   112
      Top             =   45
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Reincorporación"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":768F
      ImageAlignment  =   4
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9030
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3245
            MinWidth        =   3245
            Object.ToolTipText     =   "Usuario Ingresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3599
            MinWidth        =   3599
            Object.ToolTipText     =   "Fecha Ingreso"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3245
            MinWidth        =   3245
            Object.ToolTipText     =   "Usuario Modifica"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3599
            MinWidth        =   3599
            Object.ToolTipText     =   "Fecha Modificacion"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedAlternativa 
      Height          =   330
      Left            =   4920
      TabIndex        =   6
      Top             =   840
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7695
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
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
      PaintManager.Position=   2
      ItemCount       =   7
      Item(0).Caption =   "Datos de Contacto"
      Item(0).ControlCount=   17
      Item(0).Control(0)=   "cboSexo"
      Item(0).Control(1)=   "cboEstado"
      Item(0).Control(2)=   "dtpNacimiento"
      Item(0).Control(3)=   "txtApellido1"
      Item(0).Control(4)=   "txtApellido2"
      Item(0).Control(5)=   "Label2"
      Item(0).Control(6)=   "Label3"
      Item(0).Control(7)=   "Label4"
      Item(0).Control(8)=   "Label1(0)"
      Item(0).Control(9)=   "Label14"
      Item(0).Control(10)=   "Label15(0)"
      Item(0).Control(11)=   "txtNombre"
      Item(0).Control(12)=   "fraTipo"
      Item(0).Control(13)=   "gbPersona(0)"
      Item(0).Control(14)=   "dtpCedulaVence"
      Item(0).Control(15)=   "Label19"
      Item(0).Control(16)=   "tcAux"
      Item(1).Caption =   "Laboral"
      Item(1).ControlCount=   35
      Item(1).Control(0)=   "gbNombramiento"
      Item(1).Control(1)=   "txtInstitucionCod"
      Item(1).Control(2)=   "txtInstitucionDesc"
      Item(1).Control(3)=   "txtDeptCodigo"
      Item(1).Control(4)=   "txtDeptDesc"
      Item(1).Control(5)=   "txtSecCodigo"
      Item(1).Control(6)=   "txtSecDesc"
      Item(1).Control(7)=   "txtProfesionCod"
      Item(1).Control(8)=   "txtProfesionDesc"
      Item(1).Control(9)=   "txtSectorCod"
      Item(1).Control(10)=   "txtSectorDesc"
      Item(1).Control(11)=   "txtDeduccionesCod"
      Item(1).Control(12)=   "txtDeduccionesDesc"
      Item(1).Control(13)=   "Label10(10)"
      Item(1).Control(14)=   "Label21(0)"
      Item(1).Control(15)=   "Label9(0)"
      Item(1).Control(16)=   "lblSeccion"
      Item(1).Control(17)=   "lblDepartamento"
      Item(1).Control(18)=   "Label10(12)"
      Item(1).Control(19)=   "GroupBox1"
      Item(1).Control(20)=   "FlatScrollBarRL(0)"
      Item(1).Control(21)=   "FlatScrollBarRL(1)"
      Item(1).Control(22)=   "FlatScrollBarRL(2)"
      Item(1).Control(23)=   "FlatScrollBarRL(3)"
      Item(1).Control(24)=   "FlatScrollBarRL(4)"
      Item(1).Control(25)=   "FlatScrollBarDeduciones"
      Item(1).Control(26)=   "cboNivelAcademico"
      Item(1).Control(27)=   "Label9(1)"
      Item(1).Control(28)=   "txtCT"
      Item(1).Control(29)=   "txtCT_Desc"
      Item(1).Control(30)=   "FlatScrollBarRL(7)"
      Item(1).Control(31)=   "Label22"
      Item(1).Control(32)=   "Label9(3)"
      Item(1).Control(33)=   "txtPuestoDesc"
      Item(1).Control(34)=   "tcTrabajo"
      Item(2).Caption =   "Redes y Otros"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "GroupBox2"
      Item(3).Caption =   "Detalles"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "TituloOpciones"
      Item(3).Control(1)=   "lswHistorico"
      Item(3).Control(2)=   "btnEditarDetalle"
      Item(3).Control(3)=   "btnExport"
      Item(4).Caption =   "Bloqueos y Notas"
      Item(4).ControlCount=   14
      Item(4).Control(0)=   "cmdCambioPriDeduc"
      Item(4).Control(1)=   "chkBloqueo"
      Item(4).Control(2)=   "txtNotasAdv"
      Item(4).Control(3)=   "chkDobleDeduccion"
      Item(4).Control(4)=   "txtPriDeduc"
      Item(4).Control(5)=   "cmdNotasAdv"
      Item(4).Control(6)=   "Label16(0)"
      Item(4).Control(7)=   "Label16(3)"
      Item(4).Control(8)=   "Label20"
      Item(4).Control(9)=   "chkDesactivaAporte"
      Item(4).Control(10)=   "udPriDeduc"
      Item(4).Control(11)=   "lblOficina"
      Item(4).Control(12)=   "chkAportePatronalAdministra"
      Item(4).Control(13)=   "chkConsentimiento"
      Item(5).Caption =   "Historico"
      Item(5).ControlCount=   2
      Item(5).Control(0)=   "gbPersona(2)"
      Item(5).Control(1)=   "lswDireccion"
      Item(6).Caption =   "Cumplimiento"
      Item(6).ControlCount=   3
      Item(6).Control(0)=   "gbCProductos(0)"
      Item(6).Control(1)=   "gbCProductos(2)"
      Item(6).Control(2)=   "GroupBox4"
      Begin XtremeSuiteControls.ListView lswDireccion 
         Height          =   2415
         Left            =   -70000
         TabIndex        =   102
         Top             =   0
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   4260
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
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswHistorico 
         Height          =   6975
         Left            =   -70000
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   12303
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
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcTrabajo 
         Height          =   1335
         Left            =   -70000
         TabIndex        =   257
         Top             =   3360
         Visible         =   0   'False
         Width           =   9495
         _Version        =   1572864
         _ExtentX        =   16748
         _ExtentY        =   2355
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
         Item(0).Caption =   "Actividad Económica"
         Item(0).ControlCount=   8
         Item(0).Control(0)=   "txtTipoSociedadCod"
         Item(0).Control(1)=   "txtTipoSociedadDesc"
         Item(0).Control(2)=   "txtActividadCod"
         Item(0).Control(3)=   "txtActividadDesc"
         Item(0).Control(4)=   "FlatScrollBarRL(5)"
         Item(0).Control(5)=   "FlatScrollBarRL(6)"
         Item(0).Control(6)=   "Label10(14)"
         Item(0).Control(7)=   "Label10(13)"
         Item(1).Caption =   "Dirección de Trabajo"
         Item(1).ControlCount=   6
         Item(1).Control(0)=   "cboTraProvincia"
         Item(1).Control(1)=   "cboTraCanton"
         Item(1).Control(2)=   "cboTraDistrito"
         Item(1).Control(3)=   "txtTraDireccion"
         Item(1).Control(4)=   "btnTraDireccion"
         Item(1).Control(5)=   "Label23"
         Begin XtremeSuiteControls.ComboBox cboTraProvincia 
            Height          =   330
            Left            =   -68080
            TabIndex        =   258
            Top             =   360
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.ComboBox cboTraCanton 
            Height          =   330
            Left            =   -66160
            TabIndex        =   259
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.ComboBox cboTraDistrito 
            Height          =   330
            Left            =   -63880
            TabIndex        =   260
            Top             =   360
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1572864
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
         End
         Begin XtremeSuiteControls.FlatEdit txtTraDireccion 
            Height          =   555
            Left            =   -68080
            TabIndex        =   261
            Top             =   720
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11245
            _ExtentY        =   979
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
         Begin XtremeSuiteControls.PushButton btnTraDireccion 
            Height          =   375
            Left            =   -61600
            TabIndex        =   262
            ToolTipText     =   "Agregar Dirección"
            Top             =   840
            Visible         =   0   'False
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   79
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
            Picture         =   "frmAF_Principal.frx":7DB6
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoSociedadCod 
            Height          =   315
            Left            =   1920
            TabIndex        =   264
            Top             =   480
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
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
         Begin XtremeSuiteControls.FlatEdit txtTipoSociedadDesc 
            Height          =   315
            Left            =   2760
            TabIndex        =   265
            Top             =   480
            Width           =   5535
            _Version        =   1572864
            _ExtentX        =   9763
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtActividadCod 
            Height          =   315
            Left            =   1920
            TabIndex        =   266
            Top             =   840
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
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
         Begin XtremeSuiteControls.FlatEdit txtActividadDesc 
            Height          =   315
            Left            =   2760
            TabIndex        =   267
            Top             =   840
            Width           =   5535
            _Version        =   1572864
            _ExtentX        =   9763
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
            Height          =   255
            Index           =   5
            Left            =   8400
            TabIndex        =   268
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
            Height          =   255
            Index           =   6
            Left            =   8400
            TabIndex        =   269
            Top             =   840
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin VB.Label Label10 
            Caption         =   "Actividad Eco."
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
            Index           =   13
            Left            =   360
            TabIndex        =   271
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo Sociedad"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   360
            TabIndex        =   270
            Top             =   480
            Width           =   1455
         End
         Begin XtremeSuiteControls.Label Label23 
            Height          =   615
            Left            =   -69760
            TabIndex        =   263
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Dirección de Trabajo"
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
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   7095
         Left            =   -70000
         TabIndex        =   191
         Top             =   0
         Visible         =   0   'False
         Width           =   9375
         _Version        =   1572864
         _ExtentX        =   16536
         _ExtentY        =   12515
         _StockProps     =   79
         Caption         =   "Redes Sociales y Presencia Web"
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
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtSN_Facebook 
            Height          =   435
            Left            =   2880
            TabIndex        =   192
            Top             =   450
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
            _ExtentY        =   767
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSN_Twitter 
            Height          =   435
            Left            =   2880
            TabIndex        =   193
            Top             =   1050
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
            _ExtentY        =   767
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSN_LinkedIn 
            Height          =   435
            Left            =   2880
            TabIndex        =   194
            Top             =   1650
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
            _ExtentY        =   767
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSN_Instagram 
            Height          =   435
            Left            =   2880
            TabIndex        =   195
            Top             =   2250
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
            _ExtentY        =   767
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSN_Blog 
            Height          =   435
            Left            =   2880
            TabIndex        =   196
            Top             =   2850
            Width           =   6255
            _Version        =   1572864
            _ExtentX        =   11033
            _ExtentY        =   767
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   4
            Left            =   720
            Picture         =   "frmAF_Principal.frx":84D6
            Stretch         =   -1  'True
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   0
            Left            =   720
            Picture         =   "frmAF_Principal.frx":B400
            Stretch         =   -1  'True
            Top             =   450
            Width           =   1755
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   1
            Left            =   720
            Picture         =   "frmAF_Principal.frx":E56F
            Stretch         =   -1  'True
            Top             =   1050
            Width           =   1695
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   2
            Left            =   720
            Picture         =   "frmAF_Principal.frx":11453
            Stretch         =   -1  'True
            Top             =   1650
            Width           =   1695
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   3
            Left            =   720
            Picture         =   "frmAF_Principal.frx":13856
            Stretch         =   -1  'True
            Top             =   2250
            Width           =   1695
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCT_Desc 
         Height          =   315
         Left            =   -67240
         TabIndex        =   188
         Top             =   1560
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbCProductos 
         Height          =   2655
         Index           =   0
         Left            =   -70000
         TabIndex        =   137
         Top             =   0
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   4683
         _StockProps     =   79
         Caption         =   "Cuáles productos posee con la Organización"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.ListView lswCumplimiento 
            Height          =   2175
            Left            =   0
            TabIndex        =   143
            Top             =   240
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7646
            _ExtentY        =   3836
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
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            BackColor       =   16777215
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.GroupBox gbCProductos 
            Height          =   1815
            Index           =   1
            Left            =   4440
            TabIndex        =   139
            Top             =   0
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7646
            _ExtentY        =   3201
            _StockProps     =   79
            Caption         =   "Tipo ahorro CES"
            ForeColor       =   8421504
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
            Begin XtremeSuiteControls.RadioButton rbCES 
               Height          =   375
               Index           =   0
               Left            =   600
               TabIndex        =   140
               Top             =   360
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "[CES 1] hasta $ 1,000"
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
               Appearance      =   17
            End
            Begin XtremeSuiteControls.RadioButton rbCES 
               Height          =   375
               Index           =   1
               Left            =   600
               TabIndex        =   141
               Top             =   720
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "[CES 2] hasta $ 2,000"
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
               Appearance      =   17
            End
            Begin XtremeSuiteControls.RadioButton rbCES 
               Height          =   375
               Index           =   2
               Left            =   600
               TabIndex        =   142
               Top             =   1080
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "[CES 3] hasta $ 10,000"
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
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkBienes 
               Height          =   255
               Left            =   600
               TabIndex        =   214
               Top             =   1560
               Width           =   4215
               _Version        =   1572864
               _ExtentX        =   7435
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Posee propiedades ó bienes inmuebles"
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
               Appearance      =   17
            End
         End
         Begin XtremeSuiteControls.ComboBox cboActividad 
            Height          =   330
            Left            =   5040
            TabIndex        =   167
            Top             =   2160
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
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
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Actividad"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   2
            Left            =   5040
            TabIndex        =   168
            Top             =   1920
            Width           =   1215
         End
      End
      Begin XtremeSuiteControls.UpDown udPriDeduc 
         Height          =   315
         Left            =   -63520
         TabIndex        =   118
         Top             =   840
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   556
         _StockProps     =   64
         Appearance      =   16
         UseVisualStyle  =   0   'False
         BuddyControl    =   ""
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.CheckBox chkDesactivaAporte 
         Height          =   255
         Left            =   -68680
         TabIndex        =   66
         Top             =   2880
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9546
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Desactiva Aporte Patronal (Caso de Doble Asociación)"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin VB.Frame fraTipo 
         BorderStyle     =   0  'None
         Height          =   1290
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   8775
         Begin XtremeSuiteControls.FlatEdit txtNombreComercial 
            Height          =   315
            Left            =   1560
            TabIndex        =   19
            Top             =   120
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11245
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
         Begin XtremeSuiteControls.FlatEdit txtRazonSocial 
            Height          =   315
            Left            =   1560
            TabIndex        =   20
            Top             =   480
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11245
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
         Begin VB.Label Label15 
            Caption         =   "Nombre Comercial"
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
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Top             =   165
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Razón Social"
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
            Index           =   2
            Left            =   0
            TabIndex        =   21
            Top             =   525
            Width           =   1455
         End
      End
      Begin XtremeSuiteControls.GroupBox gbPersona 
         Height          =   1455
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   8895
         _Version        =   1572864
         _ExtentX        =   15690
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Datos de Ingreso"
         ForeColor       =   8421504
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkBeneficiarios 
            Height          =   270
            Left            =   2280
            TabIndex        =   201
            Top             =   1080
            Width           =   2415
            _Version        =   1572864
            _ExtentX        =   4260
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Desea Incluir Beneficiarios ?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaIngreso 
            Height          =   315
            Left            =   6600
            TabIndex        =   9
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   556
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
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.ComboBox cboEstadoPersona 
            Height          =   330
            Left            =   1440
            TabIndex        =   10
            Top             =   360
            Width           =   1695
            _Version        =   1572864
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
         End
         Begin XtremeSuiteControls.FlatEdit txtPromotorCod 
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   720
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
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
         Begin XtremeSuiteControls.FlatEdit txtPromotorDesc 
            Height          =   315
            Left            =   2280
            TabIndex        =   12
            Top             =   720
            Width           =   5655
            _Version        =   1572864
            _ExtentX        =   9975
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtBoleta 
            Height          =   315
            Left            =   4200
            TabIndex        =   13
            Top             =   360
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1714
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
            Text            =   "1"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarPromotor 
            Height          =   255
            Left            =   8040
            TabIndex        =   127
            Top             =   720
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin XtremeSuiteControls.FlatEdit txtHijos 
            Height          =   315
            Left            =   6960
            TabIndex        =   215
            Top             =   1080
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1714
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Dependientes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   216
            ToolTipText     =   "Número de Dependientes"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Promotor"
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
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   852
         End
         Begin VB.Label Label5 
            Caption         =   "Estado"
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
            TabIndex        =   16
            Top             =   360
            Width           =   732
         End
         Begin VB.Label Label6 
            Caption         =   "Ingreso"
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
            Left            =   5640
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Boleta"
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
            Left            =   3240
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
      End
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   330
         Left            =   1800
         TabIndex        =   23
         Top             =   1200
         Width           =   2055
         _Version        =   1572864
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
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   1800
         TabIndex        =   24
         Top             =   840
         Width           =   2055
         _Version        =   1572864
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpNacimiento 
         Height          =   315
         Left            =   6840
         TabIndex        =   25
         Top             =   840
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.GroupBox gbNombramiento 
         Height          =   1215
         Left            =   -69760
         TabIndex        =   29
         Top             =   4680
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Nombramientos"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnNombramiento 
            Height          =   495
            Left            =   7440
            TabIndex        =   75
            Top             =   720
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Agregar"
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
            Picture         =   "frmAF_Principal.frx":1690C
         End
         Begin XtremeSuiteControls.DateTimePicker dtpNombramiento 
            Height          =   315
            Left            =   1680
            TabIndex        =   30
            Top             =   720
            Width           =   1335
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
         Begin XtremeSuiteControls.FlatEdit txtAniosSerivicio 
            Height          =   315
            Left            =   3120
            TabIndex        =   31
            ToolTipText     =   "Años Laborados"
            Top             =   720
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1503
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
         Begin XtremeSuiteControls.FlatEdit txtNumeroPagos 
            Height          =   315
            Left            =   5760
            TabIndex        =   32
            Top             =   720
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
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
            Text            =   "2"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboEstadoLaboral 
            Height          =   315
            Left            =   1680
            TabIndex        =   74
            Top             =   360
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.CheckBox chkTrabajoPropio 
            Height          =   375
            Left            =   7440
            TabIndex        =   202
            Top             =   240
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Trabaja en lo propio?"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboPatrono 
            Height          =   330
            Left            =   5760
            TabIndex        =   203
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo de Patrono:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   4080
            TabIndex        =   204
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -120
            TabIndex        =   73
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "No. Pagos al Mes:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   4080
            TabIndex        =   34
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "A partir de:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   -120
            TabIndex        =   33
            Top             =   720
            Width           =   1575
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtInstitucionCod 
         Height          =   315
         Left            =   -68080
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtInstitucionDesc 
         Height          =   315
         Left            =   -67240
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
         Height          =   315
         Left            =   -68080
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
         Height          =   315
         Left            =   -67240
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
         Height          =   315
         Left            =   -68080
         TabIndex        =   39
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtSecDesc 
         Height          =   315
         Left            =   -67240
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProfesionCod 
         Height          =   315
         Left            =   -68080
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtProfesionDesc 
         Height          =   315
         Left            =   -67240
         TabIndex        =   42
         Top             =   2520
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSectorCod 
         Height          =   315
         Left            =   -68080
         TabIndex        =   43
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtSectorDesc 
         Height          =   315
         Left            =   -67240
         TabIndex        =   44
         Top             =   2880
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDeduccionesCod 
         Height          =   315
         Left            =   -68080
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtDeduccionesDesc 
         Height          =   315
         Left            =   -67240
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdCambioPriDeduc 
         Height          =   300
         Left            =   -62920
         TabIndex        =   59
         Top             =   840
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   450
         _ExtentY        =   529
         _StockProps     =   79
         BackColor       =   -2147483633
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmAF_Principal.frx":1702C
      End
      Begin XtremeSuiteControls.PushButton cmdNotasAdv 
         Height          =   300
         Left            =   -62920
         TabIndex        =   60
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   661
         _ExtentY        =   529
         _StockProps     =   79
         BackColor       =   -2147483633
         Transparent     =   -1  'True
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Principal.frx":17753
      End
      Begin XtremeSuiteControls.FlatEdit txtNotasAdv 
         Height          =   915
         Left            =   -69280
         TabIndex        =   65
         Top             =   5640
         Visible         =   0   'False
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11028
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
      Begin XtremeSuiteControls.CheckBox chkDobleDeduccion 
         Height          =   495
         Left            =   -68680
         TabIndex        =   67
         Top             =   3600
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9546
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Aplica cobro de Doble Cuota en Planillas (Caso de algunos Interinos)"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkBloqueo 
         Height          =   375
         Left            =   -68680
         TabIndex        =   68
         Top             =   4920
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9546
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Bloqueo / Desbloqueo de la Persona"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtPriDeduc 
         Height          =   315
         Left            =   -64840
         TabIndex        =   69
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.PushButton btnEditarDetalle 
         Height          =   345
         Left            =   -62440
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   614
         _StockProps     =   79
         Caption         =   "Editar"
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
         TextAlignment   =   1
         Appearance      =   7
         Picture         =   "frmAF_Principal.frx":17E7A
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1455
         Left            =   -69760
         TabIndex        =   76
         Top             =   5880
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Salario"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkSalarioEmbargos 
            Height          =   255
            Left            =   7440
            TabIndex        =   90
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Embargos?"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnSalarios 
            Height          =   495
            Left            =   7440
            TabIndex        =   77
            Top             =   840
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Agregar"
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
            Picture         =   "frmAF_Principal.frx":18475
         End
         Begin XtremeSuiteControls.DateTimePicker dtpSalarioFecha 
            Height          =   315
            Left            =   1680
            TabIndex        =   78
            Top             =   1080
            Width           =   1335
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
         Begin XtremeSuiteControls.ComboBox cboSalarioTipo 
            Height          =   315
            Left            =   1680
            TabIndex        =   80
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.ComboBox cboSalarioDivisa 
            Height          =   315
            Left            =   1680
            TabIndex        =   83
            Top             =   720
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.FlatEdit txtSalarioDevengado 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   5760
            TabIndex        =   79
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalarioRebajos 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   5760
            TabIndex        =   86
            Top             =   720
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalarioNeto 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   5760
            TabIndex        =   88
            Top             =   1080
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Salario Neto:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   4080
            TabIndex        =   89
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Rebajos:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   4080
            TabIndex        =   87
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Salario Devengado:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   4080
            TabIndex        =   85
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Divisa:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   -120
            TabIndex        =   84
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "A partir de:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   -120
            TabIndex        =   82
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -120
            TabIndex        =   81
            Top             =   360
            Width           =   1575
         End
      End
      Begin XtremeSuiteControls.GroupBox gbPersona 
         Height          =   4935
         Index           =   2
         Left            =   -70000
         TabIndex        =   91
         Top             =   2400
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   8705
         _StockProps     =   79
         Caption         =   "Datos de Localización"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.FlatEdit txtDir_Email1 
            Height          =   315
            Left            =   1920
            TabIndex        =   92
            Top             =   720
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDir_Email2 
            Height          =   315
            Left            =   1920
            TabIndex        =   93
            Top             =   1080
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDir_Direccion 
            Height          =   675
            Left            =   1920
            TabIndex        =   94
            Top             =   1920
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
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
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnDir_Eliminar 
            Height          =   495
            Left            =   6000
            TabIndex        =   104
            Top             =   3480
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Eliminar del Histórico"
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
            Picture         =   "frmAF_Principal.frx":18B95
         End
         Begin XtremeSuiteControls.FlatEdit txtDir_Id 
            Height          =   315
            Left            =   7080
            TabIndex        =   103
            Top             =   360
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDir_Distrito 
            Height          =   315
            Left            =   5760
            TabIndex        =   107
            Top             =   1560
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDir_Canton 
            Height          =   315
            Left            =   3840
            TabIndex        =   106
            Top             =   1560
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
         Begin XtremeSuiteControls.FlatEdit txtDir_Provincia 
            Height          =   315
            Left            =   1920
            TabIndex        =   105
            Top             =   1560
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
         Begin XtremeSuiteControls.FlatEdit txtDir_Telefono2 
            Height          =   315
            Left            =   6000
            TabIndex        =   99
            Top             =   2760
            Width           =   2055
            _Version        =   1572864
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDir_Telefono1 
            Height          =   315
            Left            =   1920
            TabIndex        =   98
            Top             =   2760
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
         Begin XtremeSuiteControls.FlatEdit txtDir_Usuario 
            Height          =   315
            Left            =   1920
            TabIndex        =   108
            Top             =   3240
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   550
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDir_Fecha 
            Height          =   315
            Left            =   1920
            TabIndex        =   109
            Top             =   3600
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   550
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Transparent     =   -1  'True
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha: "
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
            Index           =   15
            Left            =   360
            TabIndex        =   111
            Top             =   3600
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario: "
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
            Index           =   11
            Left            =   360
            TabIndex        =   110
            Top             =   3240
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono No.1"
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
            Index           =   8
            Left            =   360
            TabIndex        =   101
            Top             =   2760
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono No.2"
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
            Index           =   7
            Left            =   4440
            TabIndex        =   100
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Email No.1"
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
            Index           =   6
            Left            =   360
            TabIndex        =   97
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Email No.2"
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
            Index           =   1
            Left            =   360
            TabIndex        =   96
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
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
            Left            =   360
            TabIndex        =   95
            Top             =   1560
            Width           =   1095
         End
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
         Height          =   255
         Index           =   0
         Left            =   -61600
         TabIndex        =   128
         Top             =   120
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
         Height          =   255
         Index           =   1
         Left            =   -61600
         TabIndex        =   129
         Top             =   840
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
         Height          =   255
         Index           =   2
         Left            =   -61600
         TabIndex        =   130
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
         Height          =   255
         Index           =   3
         Left            =   -61600
         TabIndex        =   131
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
         Height          =   255
         Index           =   4
         Left            =   -61600
         TabIndex        =   132
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarDeduciones 
         Height          =   255
         Left            =   -61600
         TabIndex        =   133
         Top             =   480
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.ComboBox cboNivelAcademico 
         Height          =   330
         Left            =   -68080
         TabIndex        =   135
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
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
      Begin XtremeSuiteControls.GroupBox gbCProductos 
         Height          =   3495
         Index           =   2
         Left            =   -70000
         TabIndex        =   138
         Top             =   3840
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   6165
         _StockProps     =   79
         Caption         =   "Relación o Parentesco con algún empleado de la organización?   "
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.ListView lswRelacion 
            Height          =   1695
            Left            =   0
            TabIndex        =   149
            Top             =   1800
            Width           =   9255
            _Version        =   1572864
            _ExtentX        =   16325
            _ExtentY        =   2990
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
            View            =   3
            FullRowSelect   =   -1  'True
            BackColor       =   16777215
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnRelacion 
            Height          =   375
            Index           =   0
            Left            =   8280
            TabIndex        =   162
            Top             =   1320
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_Principal.frx":19139
         End
         Begin XtremeSuiteControls.RadioButton rbCRelacionParentesco 
            Height          =   255
            Index           =   0
            Left            =   4800
            TabIndex        =   145
            Top             =   0
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sí"
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
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbCRelacionParentesco 
            Height          =   255
            Index           =   1
            Left            =   5550
            TabIndex        =   146
            Top             =   0
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No"
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
            Appearance      =   17
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboR_TipoId 
            Height          =   330
            Left            =   2520
            TabIndex        =   150
            Top             =   720
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtR_Cedula 
            Height          =   330
            Left            =   4920
            TabIndex        =   152
            Top             =   720
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboR_TipoVinculo 
            Height          =   330
            Left            =   120
            TabIndex        =   154
            Top             =   720
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtR_Apellido1 
            Height          =   330
            Left            =   120
            TabIndex        =   156
            Top             =   1320
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtR_Apellido2 
            Height          =   330
            Left            =   2520
            TabIndex        =   158
            Top             =   1320
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtR_Nombre 
            Height          =   330
            Left            =   4920
            TabIndex        =   160
            Top             =   1320
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnRelacion 
            Height          =   375
            Index           =   1
            Left            =   8760
            TabIndex        =   163
            ToolTipText     =   "Elimina"
            Top             =   1320
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_Principal.frx":19859
         End
         Begin XtremeSuiteControls.PushButton btnRelacion 
            Height          =   375
            Index           =   2
            Left            =   7320
            TabIndex        =   197
            Top             =   1320
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Nuevo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_Principal.frx":19DFD
         End
         Begin XtremeSuiteControls.FlatEdit txtR_Id 
            Height          =   330
            Left            =   7200
            TabIndex        =   198
            Top             =   720
            Width           =   1095
            _Version        =   1572864
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
            Text            =   "0"
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Id. Relación"
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
            Index           =   22
            Left            =   7200
            TabIndex        =   199
            Top             =   480
            Width           =   1095
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
            Height          =   255
            Index           =   21
            Left            =   4920
            TabIndex        =   161
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido 2"
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
            Index           =   20
            Left            =   2520
            TabIndex        =   159
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido 1"
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
            Index           =   19
            Left            =   120
            TabIndex        =   157
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Relación:"
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
            Index           =   18
            Left            =   120
            TabIndex        =   155
            Top             =   480
            Width           =   1935
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
            Height          =   255
            Index           =   17
            Left            =   4920
            TabIndex        =   153
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Identificación"
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
            Index           =   16
            Left            =   2520
            TabIndex        =   151
            Top             =   480
            Width           =   1935
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCedulaVence 
         Height          =   315
         Left            =   6840
         TabIndex        =   164
         Top             =   1200
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.CheckBox chkAportePatronalAdministra 
         Height          =   255
         Left            =   -68680
         TabIndex        =   166
         Top             =   2280
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9551
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Administra el Aporte Patronal?"
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
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkConsentimiento 
         Height          =   255
         Left            =   -68680
         TabIndex        =   169
         Top             =   1800
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9551
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Firma Consentimiento Informado?"
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
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCT 
         Height          =   315
         Left            =   -68080
         TabIndex        =   187
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
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
      Begin MSComCtl2.FlatScrollBar FlatScrollBarRL 
         Height          =   255
         Index           =   7
         Left            =   -61600
         TabIndex        =   189
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtPuestoDesc 
         Height          =   330
         Left            =   -65680
         TabIndex        =   205
         Top             =   2160
         Visible         =   0   'False
         Width           =   3975
         _Version        =   1572864
         _ExtentX        =   7011
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   315
         Left            =   2760
         TabIndex        =   27
         Top             =   480
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   5280
         TabIndex        =   28
         Top             =   480
         Width           =   2895
         _Version        =   1572864
         _ExtentX        =   5101
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   1215
         Left            =   -70000
         TabIndex        =   206
         Top             =   2640
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Desempeña Cargo Político"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.CheckBox chkPersonaPolitica 
            Height          =   255
            Left            =   1200
            TabIndex        =   207
            Top             =   480
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sí"
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
         End
         Begin XtremeSuiteControls.DateTimePicker dtpC_CargoInicio 
            Height          =   330
            Left            =   2280
            TabIndex        =   208
            Top             =   480
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.DateTimePicker dtpC_CargoCorte 
            Height          =   330
            Left            =   3600
            TabIndex        =   209
            Top             =   480
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtC_CargoPolitico 
            Height          =   330
            Left            =   5040
            TabIndex        =   210
            Top             =   480
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   255
            Left            =   2280
            TabIndex        =   212
            Top             =   240
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Periodo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Cargo"
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
            Index           =   23
            Left            =   5040
            TabIndex        =   211
            Top             =   240
            Width           =   1695
         End
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   345
         Left            =   -61360
         TabIndex        =   213
         Top             =   0
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   609
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
         Transparent     =   -1  'True
         TextAlignment   =   1
         Appearance      =   7
         Picture         =   "frmAF_Principal.frx":1A42F
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   3975
         Left            =   0
         TabIndex        =   217
         Top             =   3360
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   7011
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
         Item(0).Caption =   "Localización"
         Item(0).ControlCount=   17
         Item(0).Control(0)=   "cboProvincia"
         Item(0).Control(1)=   "cboCanton"
         Item(0).Control(2)=   "cboDistrito"
         Item(0).Control(3)=   "cboNacionalidad"
         Item(0).Control(4)=   "txtEmail"
         Item(0).Control(5)=   "txtEmail_02"
         Item(0).Control(6)=   "txtApartado"
         Item(0).Control(7)=   "txtDireccion"
         Item(0).Control(8)=   "txtNotificaciones"
         Item(0).Control(9)=   "cboPaisNac"
         Item(0).Control(10)=   "Label18(10)"
         Item(0).Control(11)=   "Label18(2)"
         Item(0).Control(12)=   "Label(25)"
         Item(0).Control(13)=   "Label7"
         Item(0).Control(14)=   "Label10(9)"
         Item(0).Control(15)=   "Label11"
         Item(0).Control(16)=   "Label10(0)"
         Item(1).Caption =   "Cónyuge y Albacea"
         Item(1).ControlCount=   22
         Item(1).Control(0)=   "txtConyugeNombre"
         Item(1).Control(1)=   "txtConyugeCedula"
         Item(1).Control(2)=   "txtConyugeTelTrabajo"
         Item(1).Control(3)=   "txtConyugeTelCelular"
         Item(1).Control(4)=   "txtConyugeTelTrabajoExt"
         Item(1).Control(5)=   "Label(16)"
         Item(1).Control(6)=   "Label(15)"
         Item(1).Control(7)=   "Label(19)"
         Item(1).Control(8)=   "Label(18)"
         Item(1).Control(9)=   "Label(17)"
         Item(1).Control(10)=   "txtAlbaceaNombre"
         Item(1).Control(11)=   "txtAlbaceaCedula"
         Item(1).Control(12)=   "txtAlbaceaTelTrabajo"
         Item(1).Control(13)=   "txtAlbaceaTelTrabajoExt"
         Item(1).Control(14)=   "Label(4)"
         Item(1).Control(15)=   "Label(3)"
         Item(1).Control(16)=   "Label(2)"
         Item(1).Control(17)=   "Label(0)"
         Item(1).Control(18)=   "Label(1)"
         Item(1).Control(19)=   "txtAlbaceaTelCelular"
         Item(1).Control(20)=   "ShortcutCaption3(0)"
         Item(1).Control(21)=   "ShortcutCaption3(1)"
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   330
            Left            =   1680
            TabIndex        =   218
            Top             =   2160
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   330
            Left            =   3600
            TabIndex        =   219
            Top             =   2160
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   330
            Left            =   5880
            TabIndex        =   220
            Top             =   2160
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.ComboBox cboNacionalidad 
            Height          =   330
            Left            =   5760
            TabIndex        =   221
            Top             =   600
            Width           =   2415
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   315
            Left            =   1680
            TabIndex        =   222
            Top             =   1080
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
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
         Begin XtremeSuiteControls.FlatEdit txtEmail_02 
            Height          =   315
            Left            =   1680
            TabIndex        =   223
            Top             =   1440
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
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
         Begin XtremeSuiteControls.FlatEdit txtApartado 
            Height          =   315
            Left            =   1680
            TabIndex        =   224
            Top             =   1800
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
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
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   675
            Left            =   1680
            TabIndex        =   225
            Top             =   2520
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
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
         Begin XtremeSuiteControls.FlatEdit txtNotificaciones 
            Height          =   555
            Left            =   1680
            TabIndex        =   226
            Top             =   3240
            Width           =   6495
            _Version        =   1572864
            _ExtentX        =   11456
            _ExtentY        =   979
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
         Begin XtremeSuiteControls.ComboBox cboPaisNac 
            Height          =   330
            Left            =   1680
            TabIndex        =   227
            Top             =   600
            Width           =   2415
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtConyugeNombre 
            Height          =   315
            Left            =   -67120
            TabIndex        =   235
            Top             =   1080
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1572864
            _ExtentX        =   9340
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
         Begin XtremeSuiteControls.FlatEdit txtConyugeCedula 
            Height          =   315
            Left            =   -68920
            TabIndex        =   236
            Top             =   1080
            Visible         =   0   'False
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConyugeTelTrabajo 
            Height          =   315
            Left            =   -65080
            TabIndex        =   237
            Top             =   1680
            Visible         =   0   'False
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConyugeTelCelular 
            Height          =   315
            Left            =   -67120
            TabIndex        =   238
            Top             =   1680
            Visible         =   0   'False
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConyugeTelTrabajoExt 
            Height          =   315
            Left            =   -63040
            TabIndex        =   239
            Top             =   1680
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAlbaceaNombre 
            Height          =   315
            Left            =   -67120
            TabIndex        =   245
            Top             =   2880
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1572864
            _ExtentX        =   9340
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
         Begin XtremeSuiteControls.FlatEdit txtAlbaceaCedula 
            Height          =   315
            Left            =   -68920
            TabIndex        =   246
            Top             =   2880
            Visible         =   0   'False
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAlbaceaTelTrabajo 
            Height          =   315
            Left            =   -65080
            TabIndex        =   247
            Top             =   3480
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   556
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
            BackColor       =   16777215
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAlbaceaTelTrabajoExt 
            Height          =   315
            Left            =   -63040
            TabIndex        =   248
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   550
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
            BackColor       =   16777215
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAlbaceaTelCelular 
            Height          =   315
            Left            =   -67120
            TabIndex        =   254
            Top             =   3480
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   556
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
            BackColor       =   16777215
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   255
            Index           =   1
            Left            =   -70000
            TabIndex        =   256
            Top             =   2160
            Visible         =   0   'False
            Width           =   9495
            _Version        =   1572864
            _ExtentX        =   16748
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "Datos del Albacea General para menores de edad"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   255
            Index           =   0
            Left            =   -70000
            TabIndex        =   255
            Top             =   360
            Visible         =   0   'False
            Width           =   9495
            _Version        =   1572864
            _ExtentX        =   16748
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "Datos del Cónyuge"
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
         Begin VB.Label Label 
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
            Height          =   315
            Index           =   1
            Left            =   -67120
            TabIndex        =   253
            Top             =   2640
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label 
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
            Height          =   315
            Index           =   0
            Left            =   -68920
            TabIndex        =   252
            Top             =   2640
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. Móvil"
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
            Index           =   2
            Left            =   -67120
            TabIndex        =   251
            Top             =   3240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Extensión"
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
            Index           =   3
            Left            =   -63040
            TabIndex        =   250
            Top             =   3270
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. Trabajo"
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
            Left            =   -65080
            TabIndex        =   249
            Top             =   3270
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. Trabajo"
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
            Index           =   17
            Left            =   -65080
            TabIndex        =   244
            Top             =   1470
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Extensión"
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
            Index           =   18
            Left            =   -63040
            TabIndex        =   243
            Top             =   1470
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. Móvil"
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
            Index           =   19
            Left            =   -67120
            TabIndex        =   242
            Top             =   1440
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label 
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
            Height          =   315
            Index           =   15
            Left            =   -68920
            TabIndex        =   241
            Top             =   840
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Label 
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
            Height          =   315
            Index           =   16
            Left            =   -67120
            TabIndex        =   240
            Top             =   840
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label10 
            Caption         =   "Email No.1"
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
            Index           =   0
            Left            =   240
            TabIndex        =   234
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Apto. Postal"
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
            Left            =   240
            TabIndex        =   233
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Email No.2"
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
            Index           =   9
            Left            =   240
            TabIndex        =   232
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Dirección"
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
            Left            =   240
            TabIndex        =   231
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label 
            Caption         =   "Notificaciones:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   25
            Left            =   240
            TabIndex        =   230
            Top             =   3240
            Width           =   1365
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Nacionalidad:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   229
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   "País Nacimiento:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   228
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Puesto que desempeña"
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
         Index           =   3
         Left            =   -65680
         TabIndex        =   200
         Top             =   1920
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label22 
         Caption         =   "Centro Trabajo"
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
         Left            =   -69640
         TabIndex        =   190
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin XtremeSuiteControls.Label Label19 
         Height          =   255
         Left            =   4800
         TabIndex        =   165
         Top             =   1200
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Vencimiento de Cedula"
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
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblOficina 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina de Registro..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65200
         TabIndex        =   144
         ToolTipText     =   "Oficina de Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   4200
      End
      Begin VB.Label Label9 
         Caption         =   "Nivel Academico"
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
         Index           =   1
         Left            =   -68080
         TabIndex        =   136
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin XtremeShortcutBar.ShortcutCaption TituloOpciones 
         Height          =   360
         Left            =   -70000
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   9375
         _Version        =   1572864
         _ExtentX        =   16536
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Detalles:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
      End
      Begin VB.Label Label20 
         Caption         =   "Primer Deducción"
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
         Left            =   -68680
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueo a Créditos y Notas de Advertencia"
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
         Height          =   300
         Index           =   3
         Left            =   -69280
         TabIndex        =   62
         Top             =   4560
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cambio de Primer Deducción de Aportes"
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
         Index           =   0
         Left            =   -69280
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   5172
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.1:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   240
         TabIndex        =   58
         Top             =   240
         Width           =   2292
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.2:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   57
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5280
         TabIndex        =   56
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Nacimiento"
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
         Index           =   0
         Left            =   4800
         TabIndex        =   55
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Genero"
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
         Left            =   240
         TabIndex        =   54
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Estado Civil"
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
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label Label10 
         Caption         =   "Deductora"
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
         Index           =   10
         Left            =   -69640
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Sector"
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
         Index           =   0
         Left            =   -69640
         TabIndex        =   51
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Profesión"
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
         Index           =   0
         Left            =   -69640
         TabIndex        =   50
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblSeccion 
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
         Height          =   255
         Left            =   -69640
         TabIndex        =   49
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblDepartamento 
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
         Height          =   255
         Left            =   -69640
         TabIndex        =   48
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Institución"
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
         Index           =   12
         Left            =   -69640
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   10680
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1A599
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1A80E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1AAA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1AC2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1ADC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1AF6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1B10C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1B297
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1B422
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1B526
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1B7B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1B8C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1BB3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1BC03
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1BDA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Principal.frx":1BF3F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnIngresa 
      Height          =   330
      Index           =   2
      Left            =   9360
      TabIndex        =   113
      Top             =   45
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Ajustar"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1BFEA
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnInformes 
      Height          =   330
      Index           =   1
      Left            =   5640
      TabIndex        =   116
      Top             =   45
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Listados"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1C6DD
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnInformes 
      Height          =   330
      Index           =   0
      Left            =   4440
      TabIndex        =   117
      Top             =   45
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Boleta"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1CDE4
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   119
      ToolTipText     =   "Nuevo"
      Top             =   45
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1D4EB
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   1200
      TabIndex        =   120
      ToolTipText     =   "Editar"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1DB1D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   1560
      TabIndex        =   121
      ToolTipText     =   "Eliminar"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1E118
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   2040
      TabIndex        =   122
      ToolTipText     =   "Guardar"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1E6BC
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   2400
      TabIndex        =   123
      ToolTipText     =   "Deshacer"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1EDED
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   2880
      TabIndex        =   124
      ToolTipText     =   "Reporte"
      Top             =   45
      Visible         =   0   'False
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1F4ED
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   6
      Left            =   3240
      TabIndex        =   125
      ToolTipText     =   "Consultas"
      Top             =   45
      Visible         =   0   'False
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Principal.frx":1FBF4
      ImageAlignment  =   6
   End
   Begin MSComCtl2.FlatScrollBar scrBar_Persona 
      Height          =   255
      Left            =   9360
      TabIndex        =   126
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   10560
      TabIndex        =   134
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   45
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Adjuntos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Picture         =   "frmAF_Principal.frx":202F4
   End
   Begin XtremeSuiteControls.FlatEdit txtDIMEX 
      Height          =   330
      Left            =   7080
      TabIndex        =   147
      Top             =   840
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
   Begin VB.Label lblDimex 
      BackStyle       =   0  'Transparent
      Caption         =   "DIMEX"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7080
      TabIndex        =   148
      Top             =   600
      Width           =   1695
   End
   Begin XtremeShortcutBar.ShortcutCaption TituloOpcion 
      Height          =   375
      Left            =   2760
      TabIndex        =   115
      Top             =   1320
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Datos de Contacto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   114
      Top             =   1320
      Width           =   2775
      _Version        =   1572864
      _ExtentX        =   4890
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Detalle:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Id. Alterna"
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
      Index           =   5
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1695
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
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Identificación"
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
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmAF_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type AfCambios
  vEstado As String
  vPromotor As String
  vFecNac As String
End Type

Dim vCambios As AfCambios
Dim vEditar As Boolean, vFechaActual As Date
Dim vCedula As String, vSeek As Integer, vScroll As Boolean, vTipoJuridica As Integer
Dim vPaso As Boolean

Dim vRA_Access As Boolean

Const Id_TaskItem_DatosPersonales = 0
Const Id_TaskItem_RelacionLaboral = 1
Const Id_TaskItem_Redes = 2

Const Id_TaskItem_Telefonos = 3
Const Id_TaskItem_Beneficiarios = 4
Const Id_TaskItem_Cuentas = 5
Const Id_TaskItem_Nombramientos = 6
Const Id_TaskItem_Ingresos = 7
Const Id_TaskItem_Renuncias = 8
Const Id_TaskItem_Liquidaciones = 9
Const Id_TaskItem_Bloqueos = 10
Const Id_TaskItem_Tarjetas = 11
Const Id_TaskItem_Canales = 12
Const Id_TaskItem_Preferencias = 13
Const Id_TaskItem_Bienes = 14
Const Id_TaskItem_Escolaridad = 15
Const Id_TaskItem_Direcciones = 16
Const Id_TaskItem_Salarios = 17
Const Id_TaskItem_Motivos = 18

Const Id_TaskItem_Cumplimiento = 19
Const Id_TaskItem_Emails = 20

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim mSavePass As Boolean


Private Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR", "EDICION"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub sbTaskPanel_Load()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem

    tpMain.VisualTheme = xtpTaskPanelThemeOffice2016
  
    Set Group = tpMain.Groups.Add(0, "Registro")
    Group.ToolTip = "Información Principal para el Registro de la Persona"
    Group.Special = True

    
    Group.Items.Add Id_TaskItem_DatosPersonales, "Datos Personales", xtpTaskItemTypeLink, 4
    Group.Items.Add Id_TaskItem_RelacionLaboral, "Relación Laboral", xtpTaskItemTypeLink, 1
    Group.Items.Add Id_TaskItem_Cumplimiento, "Oficina Cumplimiento", xtpTaskItemTypeLink, 10
    Group.Items.Add Id_TaskItem_Redes, "Redes Sociales", xtpTaskItemTypeLink, 10
    
    Set Group = tpMain.Groups.Add(0, "Detalles")
    Group.ToolTip = "Datos adicionales de la persona"
    
    Group.Items.Add Id_TaskItem_Telefonos, "Teléfonos", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Beneficiarios, "Beneficiarios", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Cuentas, "Cuentas Bancarias", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Tarjetas, "Tarjetas", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Direcciones, "Localización", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Bloqueos, "Bloqueos y Otros", xtpTaskItemTypeLink, 7
    
    
    Set Group = tpMain.Groups.Add(0, "Histórico")
    Group.Expanded = False
    Group.Items.Add Id_TaskItem_Ingresos, "Ingresos", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Renuncias, "Renuncias", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Liquidaciones, "Liquidaciones", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Nombramientos, "Nombramientos", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Salarios, "Salarios", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Emails, "Emails", xtpTaskItemTypeLink, 3
    
    
    
    Set Group = tpMain.Groups.Add(0, "Info adicional")
    Group.Items.Add Id_TaskItem_Motivos, "Motivos", xtpTaskItemTypeLink, 12
    Group.Items.Add Id_TaskItem_Canales, "Canales", xtpTaskItemTypeLink, 12
    Group.Items.Add Id_TaskItem_Preferencias, "Preferencias", xtpTaskItemTypeLink, 12
    Group.Items.Add Id_TaskItem_Bienes, "Bienes", xtpTaskItemTypeLink, 12
    Group.Items.Add Id_TaskItem_Escolaridad, "Escolaridad", xtpTaskItemTypeLink, 12
    
   
    tpMain.SetImageList imlTaskPanelIcons
    
'    tpMain.SetMargins 5, 5, 5, 5, 5
    

End Sub


Private Sub sbCpl_Productos()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Persona_Productos_Consulta '" & txtCedula.Text & "', 1"
Call OpenRecordSet(rs, strSQL)

vPaso = True

With lswCumplimiento
    .ListItems.Clear
    
    Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , rs!Descripcion)
            itmX.Tag = rs!Codigo
            If rs!asignado = 1 Then
              itmX.Checked = True
            End If
      rs.MoveNext
    Loop

End With
rs.Close

vPaso = False

Me.MousePointer = vbDefault


Exit Sub

vError:
  vPaso = False
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCpl_Relaciones()

On Error GoTo vError

Me.MousePointer = vbHourglass

'Limpia
Call btnRelacion_Click(2)

strSQL = "exec spAFI_Persona_Relacion_List '" & txtCedula.Text & "', 1"
Call OpenRecordSet(rs, strSQL)

With lswRelacion.ColumnHeaders
    .Clear
    .Add , , "Id Rel.", 1100, vbCenter
    .Add , , "Relación", 2100, vbCenter
    .Add , , "Cédula", 1800, vbCenter
    .Add , , "Nombre", 3800
    .Add , , "Tel.Movil", 1200
    .Add , , "Tel.Trabajo", 1200
    .Add , , "Tel.Tra.Ext", 1200
    .Add , , "Reg.Fecha", 1200
    .Add , , "Reg.Usuario", 1200, vbCenter
    .Add , , "Act.Fecha", 1200
    .Add , , "Act.Usuario", 1200, vbCenter
End With


With lswRelacion
    .ListItems.Clear
    Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , rs!Pr_Id)
            itmX.SubItems(1) = rs!Tipo_Relacion_Desc & ""
            itmX.SubItems(2) = rs!Cedula & ""
            itmX.SubItems(3) = rs!NOMBRE_COMPLETO & ""
            itmX.SubItems(4) = rs!TelCell & ""
            itmX.SubItems(5) = rs!TelTra & ""
            itmX.SubItems(6) = rs!TelTraExt & ""
            itmX.SubItems(7) = rs!Registro_Fecha & ""
            itmX.SubItems(8) = rs!Registro_Usuario & ""
            itmX.SubItems(9) = rs!MODIFICA_FECHA & ""
            itmX.SubItems(10) = rs!MODIFICA_USUARIO & ""
            
    
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

Private Sub sbCumplimiento_Load()

On Error GoTo vError

Call sbCpl_Productos
Call sbCpl_Relaciones

vError:

End Sub


Private Sub sbTaskPanel_Accion(ItemId As Integer)

Dim fraX As Frame

If Trim(txtCedula.Text) = "" Then Exit Sub

On Error GoTo vError


Select Case ItemId
  Case Id_TaskItem_DatosPersonales  'Datos de Contato
    

    TituloOpcion.Caption = "Datos de Contacto"
    
    tcMain.Item(0).Selected = True
    Call cboTipoId_Click
    
    If fraTipo.Visible Then
       txtNombreComercial.SetFocus
    Else
       If txtApellido1.Enabled Then txtApellido1.SetFocus
    End If
    
    Exit Sub
    
  Case Id_TaskItem_RelacionLaboral  'Relación Laboral
    TituloOpcion.Caption = "Datos Laborales"
    
    tcMain.Item(1).Selected = True
    tcTrabajo.Item(0).Selected = True
    
    txtInstitucionCod.SetFocus

    Exit Sub

  Case Id_TaskItem_Redes 'Información Adicional
    
    TituloOpcion.Caption = "Redes y Otros..."
    tcMain.Item(2).Selected = True
    
    Exit Sub
      
  Case Id_TaskItem_Cumplimiento   'Oficina de Cumplimiento
    

    TituloOpcion.Caption = "Datos para Oficina de Cumplimiento"
    
    tcMain.Item(6).Selected = True
    Call sbCumplimiento_Load
    
    Exit Sub
      
      
End Select



Select Case ItemId
    Case Id_TaskItem_Telefonos, Id_TaskItem_Beneficiarios, Id_TaskItem_Cuentas
     'Permite el registro de detalles sin maestro
Case Else
        If Not vEditar Then
            MsgBox "Se encuentra en modo de Registro, guarde los datos de la persona y luego ingrese a esta opción!", vbInformation
            Exit Sub
        End If
End Select

'Otras Opciones
TituloOpcion.Caption = "Otros datos:"


Select Case ItemId
    Case Id_TaskItem_Bloqueos
       tcMain.Item(4).Selected = True
    Case Else
       tcMain.Item(3).Selected = True
End Select

lswHistorico.ColumnHeaders.Clear
lswHistorico.ListItems.Clear
lswHistorico.Checkboxes = False

btnEditarDetalle.Visible = False


Select Case ItemId
  Case Id_TaskItem_Emails  'Emails
        
    TituloOpciones.Caption = "Lista de Emails..:"
    TituloOpciones.Tag = "Emails"
    
    btnEditarDetalle.Visible = False
        
    lswHistorico.ColumnHeaders.Add , , "Tipo", 2100
    lswHistorico.ColumnHeaders.Add , , "Email", 3500
    lswHistorico.ColumnHeaders.Add , , "Principal", 1500, vbCenter
    
    
    
    strSQL = "exec spAFI_Persona_Email_Consulta '" & Trim(txtCedula.Text) & "', '" & glogon.Usuario & "'"
       
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , Trim(rs!TipoDesc))
           itmX.SubItems(1) = rs!Email & ""
           itmX.SubItems(2) = IIf(rs!Principal = True, "Sí", "No")
       rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Telefonos  'Telefonos
        
    TituloOpciones.Caption = "Lista de Teléfonos..:"
    TituloOpciones.Tag = "Telefonos"
    
    btnEditarDetalle.Visible = True
        
    lswHistorico.ColumnHeaders.Add 1, , "Numero", 1500
    lswHistorico.ColumnHeaders.Add 2, , "Tipo", 1500
    lswHistorico.ColumnHeaders.Add 3, , "Extension", 1500
    lswHistorico.ColumnHeaders.Add 4, , "Contacto", 2500
    
    
    
    strSQL = "select T.Telefono, T.Tipo, T.Numero, T.Ext, T.Contacto, T.Usuario, T.Fecha, Tt.NombreTipoTelefono as 'TipoDesc'" _
           & " , dbo.MyGetdate() as FechaServidor" _
           & " from Telefonos T inner join AFI_TIPOS_TELEFONOS Tt on T.Tipo = Tt.IdTipoTelefono " _
           & " where cedula = '" & Trim(txtCedula.Text) & "'"
       
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , Trim(rs!Numero))
           itmX.SubItems(1) = rs!TipoDesc
           itmX.SubItems(2) = Trim(rs!Ext) & ""
           itmX.SubItems(3) = Trim(rs!contacto) & ""
       rs.MoveNext
    Loop
    rs.Close
    
    
  Case Id_TaskItem_Beneficiarios 'Beneficiarios
    btnEditarDetalle.Visible = True
 
     
    TituloOpciones.Caption = "Lista de Beneficiarios..:"
    TituloOpciones.Tag = "Beneficiarios"
    
    lswHistorico.ColumnHeaders.Add 1, , "Identificación", 1500
    lswHistorico.ColumnHeaders.Add 2, , "Nombre", 3500
    lswHistorico.ColumnHeaders.Add 3, , "Porcentaje", 1100, vbRightJustify
    lswHistorico.ColumnHeaders.Add 4, , "Relación", 1200, vbCenter
    lswHistorico.ColumnHeaders.Add 5, , "Parentesco", 1100, vbCenter
    

    
    strSQL = "exec spAFI_PERSONA_BENEFICIARIOS_Consulta '" & Trim(txtCedula) & "',0"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!cedula_Beneficiario)
           itmX.SubItems(1) = Trim(rs!Nombre)
           itmX.SubItems(2) = Trim(rs!Porcentaje)
           itmX.SubItems(3) = Trim(rs!Relacion_Desc)
           itmX.SubItems(4) = Trim(rs!parentesco)
       
       rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Cuentas 'Cuentas Bancarias
    
    btnEditarDetalle.Visible = True
    
    
    TituloOpciones.Caption = "Cuentas bancarias..:"
    TituloOpciones.Tag = "Cuentas"
    
    lswHistorico.ColumnHeaders.Add 1, , "Cuenta", 2500
    lswHistorico.ColumnHeaders.Add 2, , "Banco", 3500
    lswHistorico.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 8, , "Fecha", 2500
    lswHistorico.ColumnHeaders.Add 9, , "Usuario", 2500

        strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
               & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
               & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
               & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
               & " where C.Identificacion = '" & Trim(txtCedula) & "'" 'and C.Modulo = 'AFI'
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!ACTIVA = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           itmX.SubItems(8) = rs!Registro_Usuario & ""
     
       rs.MoveNext
    Loop
    rs.Close
    
    
    
  
  Case Id_TaskItem_Nombramientos 'Nombramientos
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "Estado", 1200
    lswHistorico.ColumnHeaders.Add 2, , "A Partir", 1500
    lswHistorico.ColumnHeaders.Add 3, , "Fecha", 2500
    lswHistorico.ColumnHeaders.Add 4, , "Usuario", 2500
            
    TituloOpciones.Caption = "Nombramientos..:"
    TituloOpciones.Tag = "Nombramientos"
    
    strSQL = "exec spAFI_Persona_Nombramientos_Consulta '" & txtCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lswHistorico.ListItems.Add(, , rs!EstadoLaboralDesc)
          itmX.SubItems(1) = Format(rs!fecha, "dd/mm/yyyy")
          itmX.SubItems(2) = rs!Registro_Fecha
          itmX.SubItems(3) = rs!Registro_Usuario
      rs.MoveNext
    Loop
    rs.Close
  
  
  
  Case Id_TaskItem_Salarios 'Salarios
     
    TituloOpciones.Caption = "Lista de Salarios..:"
    TituloOpciones.Tag = "Salarios"
    
    lswHistorico.ColumnHeaders.Add 1, , "Tipo", 1500
    lswHistorico.ColumnHeaders.Add 2, , "Fecha", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add 3, , "Divisa", 1000, vbCenter
    lswHistorico.ColumnHeaders.Add 4, , "Devengado", 1600, vbRightJustify
    lswHistorico.ColumnHeaders.Add 5, , "Rebajos", 1600, vbRightJustify
    lswHistorico.ColumnHeaders.Add 6, , "Neto", 1600, vbRightJustify
    lswHistorico.ColumnHeaders.Add 7, , "Embargo?", 1600, vbCenter
    lswHistorico.ColumnHeaders.Add 8, , "Fecha", 2500
    lswHistorico.ColumnHeaders.Add 9, , "Usuario", 2500

    
    strSQL = "exec spAFI_PERSONA_SALARIOS_Consulta '" & Trim(txtCedula) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!TipoSalarioDesc)
           itmX.SubItems(1) = Format(rs!Fecha_Salario, "dd/MM/yyyy")
           itmX.SubItems(2) = Trim(rs!cod_Divisa)
           itmX.SubItems(3) = Format(rs!SALARIO_DEVENGADO, "Standard")
           itmX.SubItems(4) = Format(rs!Rebajos_Total, "Standard")
           itmX.SubItems(5) = Format(rs!Salario_Neto, "Standard")
           itmX.SubItems(6) = Trim(rs!Embargo)
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           itmX.SubItems(8) = Trim(rs!Registro_Usuario & "")

       
       rs.MoveNext
    Loop
    rs.Close
  
  
  
  Case Id_TaskItem_Direcciones 'Historico de Direcciones
    btnEditarDetalle.Visible = True
 
     
    TituloOpciones.Caption = "Histórico de Direcciones y Contacto..:"
    TituloOpciones.Tag = "Direcciones"
    
    lswHistorico.ColumnHeaders.Add , , "Tipo", 1500
    lswHistorico.ColumnHeaders.Add , , "Fecha", 2500
    lswHistorico.ColumnHeaders.Add , , "Provincia", 1800
    lswHistorico.ColumnHeaders.Add , , "Canton", 1800
    lswHistorico.ColumnHeaders.Add , , "Distrito", 1800
    lswHistorico.ColumnHeaders.Add , , "Dirección", 1800
    lswHistorico.ColumnHeaders.Add , , "Email 01", 2800
    lswHistorico.ColumnHeaders.Add , , "Email 02", 2800
    lswHistorico.ColumnHeaders.Add , , "Tel No 1", 1800
    lswHistorico.ColumnHeaders.Add , , "Tel No 2", 1800

    
    strSQL = "exec spAFI_PERSONA_DIRECCIONES_Consulta '" & Trim(txtCedula.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!TipoDesc)
           itmX.SubItems(1) = Format(rs!Registro_Fecha, "yyyy-mm-dd")
           itmX.SubItems(2) = Trim(rs!ProvinciaDesc)
           itmX.SubItems(3) = Trim(rs!CantonDesc)
           itmX.SubItems(4) = Trim(rs!DistritoDesc)
           itmX.SubItems(5) = Trim(rs!direccion)
           itmX.SubItems(6) = Trim(rs!Email_01)
           itmX.SubItems(7) = Trim(rs!Email_02)
           itmX.SubItems(8) = Trim(rs!Telefono_01)
           itmX.SubItems(9) = Trim(rs!Telefono_02)
       rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Ingresos 'Ingresos
  
    With lswHistorico
            .ListItems.Clear
            .ColumnHeaders.Clear
            .ColumnHeaders.Add 1, , "ID", 900
            .ColumnHeaders.Add 2, , "Usuario", 1500
            .ColumnHeaders.Add 3, , "Fecha", 1500
            .ColumnHeaders.Add 4, , "Ingreso", 1200
            .ColumnHeaders.Add 5, , "Boleta", 1100
            .ColumnHeaders.Add 6, , "Promotor", 3500
            .ColumnHeaders.Add 7, , "Tipo Ingreso", 2500
            
            TituloOpciones.Caption = "Ingresos..:"
            TituloOpciones.Tag = "Ingresos"
            
            strSQL = "exec spAFI_Ingresos_Consulta '" & Trim(txtCedula) & "'"
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
               Set itmX = .ListItems.Add(, , rs!consec)
                   itmX.SubItems(1) = rs!Usuario & ""
                   itmX.SubItems(2) = rs!fecha & ""
                   itmX.SubItems(3) = Format(rs!Fecha_Ingreso, "yyyy-mm-dd")
                   itmX.SubItems(4) = rs!Boleta & ""
                   itmX.SubItems(5) = rs!promotor & ""
                   itmX.SubItems(6) = rs!Tipo_Desc & ""
               rs.MoveNext
            Loop
            rs.Close
    End With
  
  
  Case Id_TaskItem_Renuncias  'Renuncias
    
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "Fecha", 1500
    lswHistorico.ColumnHeaders.Add 2, , "Causa", 2500
    lswHistorico.ColumnHeaders.Add 3, , "Tipo", 1500
    
    
    TituloOpciones.Caption = "Renuncias..:"
    TituloOpciones.Tag = "Renuncias"
    
    
    strSQL = "Select R.Fecha,R.Tipo,C.Descripcion" _
           & " From Renuncias R inner join Causas_Renuncias C " _
           & " on R.id_causa=C.id_causa where R.Cedula='" & Trim(txtCedula) & "'"
           
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , Format(rs!fecha, "dd/mm/yyyy"))
           itmX.SubItems(1) = Trim(rs!Descripcion)
           itmX.SubItems(2) = Trim(rs!Tipo)
       rs.MoveNext
    Loop
    rs.Close
    
    lswHistorico.Enabled = True
  
  Case Id_TaskItem_Liquidaciones 'Liquidaciones
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "# Liquidacion", 2500
    lswHistorico.ColumnHeaders.Add 2, , "Fecha", 1500
    lswHistorico.ColumnHeaders.Add 3, , "Tipo", 1500
    
    TituloOpciones.Caption = "Liquidaciones..:"
    TituloOpciones.Tag = "Liquidaciones"
    
    strSQL = "Select Consec,Fecliq,EstadoActliq From liquidacion " _
           & "where Cedula='" & Trim(txtCedula) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!consec)
           itmX.SubItems(1) = Format(rs!fecLiq, "dd/mm/yyyy")
           itmX.SubItems(2) = rs!estadoactliq
       rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Tarjetas  'Tarjetas
  
    btnEditarDetalle.Visible = True
    
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "No. Tarjeta", 2500
    lswHistorico.ColumnHeaders.Add 2, , "Tipo", 2100
    lswHistorico.ColumnHeaders.Add 3, , "Vence", 1500
    
    TituloOpciones.Caption = "Tarjetas..:"
    TituloOpciones.Tag = "Tarjetas"
  
  
            strSQL = "exec spAFI_PersonaTarjetas_Consulta " & gPortal.Empresa_Id & ",'" & txtCedula.Text & "',''"
            Call OpenRecordSet(rs, strSQL)
            
            With lswHistorico.ListItems
               .Clear
               Do While Not rs.EOF
                Set itmX = .Add(, , rs!Tarjeta_Mask)
                    itmX.SubItems(1) = rs!Tarjeta_Tipo
                    itmX.SubItems(2) = Format(rs!Tarjeta_Vence, "MM/YY")
                rs.MoveNext
               Loop
               rs.Close
            End With



  Case Id_TaskItem_Motivos  'Motivos de Afiliación
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "Motivo de Ingreso", 3500
    lswHistorico.ColumnHeaders.Add 2, , "Usuario", 2100
    lswHistorico.ColumnHeaders.Add 3, , "Fecha", 2500
    lswHistorico.Checkboxes = True
    
    TituloOpciones.Caption = "Motivos de Afiliación..:"
    TituloOpciones.Tag = "Motivo"

  
            strSQL = "exec spAFI_Persona_Motivos_Consulta '" & txtCedula.Text & "',1"
            Call OpenRecordSet(rs, strSQL)
            
            vPaso = True
            With lswHistorico.ListItems
               .Clear
               Do While Not rs.EOF
                Set itmX = .Add(, , rs!Descripcion)
                    itmX.SubItems(1) = rs!Registro_Usuario & ""
                    itmX.SubItems(2) = rs!Registro_Fecha & ""
                    itmX.Tag = rs!Cod_Motivo
                    
                    itmX.Checked = IIf((rs!asignado = 1), True, False)
                    
                rs.MoveNext
               Loop
               rs.Close
            End With
            vPaso = False


  Case Id_TaskItem_Canales  'Canal de Comunicacion
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "Tipo de Canal", 3500
    lswHistorico.ColumnHeaders.Add 2, , "Usuario", 2100
    lswHistorico.ColumnHeaders.Add 3, , "Fecha", 2500
    lswHistorico.Checkboxes = True
    
    TituloOpciones.Caption = "Canal de Conctato..:"
    TituloOpciones.Tag = "Canal"
  'TODO
  
            strSQL = "exec spAFI_Persona_Canales_Consulta '" & txtCedula.Text & "',1"
            Call OpenRecordSet(rs, strSQL)
            
            vPaso = True
            With lswHistorico.ListItems
               .Clear
               Do While Not rs.EOF
                Set itmX = .Add(, , rs!Descripcion)
                    itmX.SubItems(1) = rs!Registro_Usuario & ""
                    itmX.SubItems(2) = rs!Registro_Fecha & ""
                    itmX.Tag = rs!Canal_Tipo
                    
                    itmX.Checked = IIf((rs!asignado = 1), True, False)
                    
                rs.MoveNext
               Loop
               rs.Close
            End With
            vPaso = False

  
  Case Id_TaskItem_Preferencias 'Gustos y Preferencias
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "Gustos y Preferencias", 3500
    lswHistorico.ColumnHeaders.Add 2, , "Usuario", 2100
    lswHistorico.ColumnHeaders.Add 3, , "Fecha", 2500
    lswHistorico.Checkboxes = True
    
    TituloOpciones.Caption = "Gustos y Preferencias..:"
    TituloOpciones.Tag = "Preferencias"
  
    strSQL = "exec spAFI_Persona_Preferencias_Consulta '" & txtCedula.Text & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswHistorico.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
            itmX.SubItems(1) = rs!Registro_Usuario & ""
            itmX.SubItems(2) = rs!Registro_Fecha & ""
            itmX.Tag = rs!cod_preferencia
            
            itmX.Checked = IIf((rs!asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
  
  
  Case Id_TaskItem_Bienes 'Bienes
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "Bienes", 3500
    lswHistorico.ColumnHeaders.Add 2, , "Usuario", 2100
    lswHistorico.ColumnHeaders.Add 3, , "Fecha", 2500
    lswHistorico.Checkboxes = True
    
    TituloOpciones.Caption = "Bienes de la Persona..:"
    TituloOpciones.Tag = "Bienes"
  
    strSQL = "exec spAFI_Persona_Bienes_Consulta '" & txtCedula.Text & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswHistorico.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
            itmX.SubItems(1) = rs!Registro_Usuario & ""
            itmX.SubItems(2) = rs!Registro_Fecha & ""
            itmX.Tag = rs!Bien_Tipo
            
            itmX.Checked = IIf((rs!asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
  
  Case Id_TaskItem_Escolaridad 'Escolaridad
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "Nivel de Escolaridad", 3500
    lswHistorico.ColumnHeaders.Add 2, , "Usuario", 2100
    lswHistorico.ColumnHeaders.Add 3, , "Fecha", 2500
    lswHistorico.Checkboxes = True
    
    TituloOpciones.Caption = "Nivel de Escolaridad..:"
    TituloOpciones.Tag = "Escolaridad"
  
    strSQL = "exec spAFI_Persona_Escolaridad_Consulta '" & txtCedula.Text & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswHistorico.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
            itmX.SubItems(1) = rs!Registro_Usuario & ""
            itmX.SubItems(2) = rs!Registro_Fecha & ""
            itmX.Tag = rs!Escolaridad_Tipo
            
            itmX.Checked = IIf((rs!asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
  
  
  
  
  Case Id_TaskItem_Bloqueos 'Bloqueos
    tcMain.Item(4).Selected = True
    
End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbClearControles()
Dim vControl As Control

If vEditar Then Exit Sub

For Each vControl In Me
  If TypeOf vControl Is TextBox Then
     vControl.Text = ""
  End If
  
  If TypeOf vControl Is XtremeSuiteControls.FlatEdit Then
     vControl.Text = ""
  End If
  
Next

StatusBarX.Panels.Item(1) = ""
StatusBarX.Panels.Item(2) = ""
StatusBarX.Panels.Item(3) = ""
StatusBarX.Panels.Item(4) = ""


End Sub

Public Sub sbConsultaExterna(pCedula As String)

On Error GoTo vError
 
 Call TimerX_Timer
 txtCedula.Text = pCedula
 txtCedAlternativa.SetFocus
 Call sbCurrentRecord

Exit Sub

vError:

End Sub

Private Sub sbCurrentRecord()
Dim rsTemp As New ADODB.Recordset
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim i As Integer, vEspacio As Integer

On Error Resume Next

If Not fxSIFValidaCadena(txtCedula.Text) Then
   Exit Sub
End If

'Valida Acceso a Expediente
vRA_Access = fxSys_RA_Consulta(Trim(txtCedula.Text), glogon.Usuario)
 
If Not vRA_Access Then
    MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
    txtCedula.Text = ""
    txtNombre.Text = ""
    Exit Sub
End If
    

strSQL = "exec spAFI_Persona_Consulta '" & Trim(txtCedula.Text) & "'"

Call OpenRecordSet(rs, strSQL)
If (Not rs.EOF And Not rs.BOF) And Not glogon.error Then
   vEditar = True
   
   
   Call sbBarra_Accion("activo")
   Call RefrescaTags(Me)
   
   Call sbLockControles("L")
   Call sbLimpiaDatos 'Inicializa Datos
   
   vCedula = Trim(rs!Cedula)
   txtCedula.Text = Trim(rs!Cedula)
   
   If Not IsNull(rs!TipoIdDesc) Then
       vPaso = True
           cboTipoId.Text = Trim(rs!TipoIdDesc)
       vPaso = False
   End If
   
   If Not IsNull(rs!Nacionalidad) Then
       vPaso = True
       Call sbCboAsignaDato(cboNacionalidad, Trim(rs!Nacionalidad), True, rs!Cod_Nacionalidad)
       vPaso = False
   End If
   
   If Not IsNull(rs!Pais) Then
       vPaso = True
       Call sbCboAsignaDato(cboPaisNac, Trim(rs!Pais), True, rs!Cod_Pais_Nac)
       vPaso = False
   End If
   

   If rs!NivelAcademicoDesc <> "" Then
       vPaso = True
       Call sbCboAsignaDato(cboNivelAcademico, Trim(rs!NivelAcademicoDesc), True, rs!Nivel_Academico)
       vPaso = False
   End If

   If rs!C_ActividadDesc <> "" Then
       vPaso = True
       Call sbCboAsignaDato(cboActividad, Trim(rs!C_ActividadDesc), True, rs!Actividades)
       vPaso = False
   End If

     
   If rs!Tipo_Personeria = "J" Then
      fraTipo.Visible = True
      txtNombreComercial.Text = Trim(rs!Nombre)
      txtRazonSocial.Text = Trim(rs!Razon_Social & "")
      
   Else
      fraTipo.Visible = False
        
        vEspacio = 1
        For i = 1 To Len(Trim(rs!Nombre))
          If Mid(Trim(rs!Nombre), i, 1) <> " " Then
             Select Case vEspacio
              Case 1
               vApellido1 = vApellido1 & Mid(Trim(rs!Nombre), i, 1)
              Case 2
               vApellido2 = vApellido2 & Mid(Trim(rs!Nombre), i, 1)
              Case 3
               vNombre1 = vNombre1 & Mid(Trim(rs!Nombre), i, 1)
              Case Is >= 4
               vNombre2 = vNombre2 & Mid(Trim(rs!Nombre), i, 1)
             End Select
          Else
             vEspacio = vEspacio + 1
          End If
        Next i
   
        txtApellido1 = vApellido1
        txtApellido2 = vApellido2
        txtNombre = vNombre1 & " " & vNombre2
   
   End If
    
     
   txtCedAlternativa = Trim(rs!cedular & "")
   
   txtActividadCod.Text = Trim(rs!Cod_actividad & "")
   txtActividadDesc.Text = rs!ActividadDesc & ""
    
   txtTipoSociedadCod.Text = Trim(rs!Cod_Sociedad & "")
   txtTipoSociedadDesc.Text = rs!SociedadDesc & ""
     
   txtBoleta = rs!id_Boleta_AF & ""
     
   If Not IsNull(rs!FECHA_VEN_CED) Then
        dtpCedulaVence.Value = rs!FECHA_VEN_CED
   End If
   
   'Carga Información del Estado de la Persona y sus posibles Acciones
   
   
   btnIngresa.Item(0).Enabled = False 'Reingreso
   strSQL = "select COD_MOVIMIENTO from AFI_ESTADOS_CAMBIO" _
          & " where COD_ESTADO = '" & rs!EstadoActual & "' and COD_MOVIMIENTO IN('REI','ACT')"
   Call OpenRecordSet(rsTemp, strSQL, 0)
   Do While Not rsTemp.EOF
    If rsTemp!COD_MOVIMIENTO = "REI" Then btnIngresa.Item(0).Enabled = True
    If rsTemp!COD_MOVIMIENTO = "ACT" Then btnIngresa.Item(0).Enabled = True
    rsTemp.MoveNext
   Loop
   rsTemp.Close
   
   'Si el Estado es de Ingreso, Puede usarse con otros estados de ingreso, caso contrario limpiar la lista
   strSQL = "select count(*) as Ingreso from AFI_ESTADOS_CAMBIO" _
          & " where COD_ESTADO = '" & rs!EstadoActual & "' and COD_MOVIMIENTO IN('ING')"
   Call OpenRecordSet(rsTemp, strSQL, 0)
   If rsTemp!Ingreso = 0 Then
       cboEstadoPersona.Clear
   End If
   rsTemp.Close
   
   Call sbCboAsignaDato(cboEstadoPersona, rs!EstadoPersonaDesc, True, rs!EstadoActual)
   
   
   If IsNull(rs!PriDeduc) Then
      txtPrideduc.Text = GLOBALES.glngFechaCR
   Else
      txtPrideduc.Text = rs!PriDeduc
   End If
     
   dtpFechaIngreso = rs!FechaIngreso
   dtpNacimiento = rs!fecha_nac
   cboSexo = IIf(rs!sexo = "M", "Masculino", "Femenino")
     
     
   Call sbCboAsignaDato(cboEstado, rs!EstadoCivilDesc & "", True, rs!EstadoCivil) 'Se activa el Click ->    Call cboProvincia_Click
     
   'Direccion Personal y Trabajo (Opcional)
     
   Call sbCboAsignaDato(cboProvincia, rs!ProvinciaDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
   Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
   Call sbCboAsignaDato(cboDistrito, rs!DistritoDesc & "")
   
   cboDistrito.ToolTipText = Trim(rs!Distrito) & ""
   txtDireccion = Trim(rs!direccion) & ""
     
  'Trabajo es Opcional
  If rs!Tra_Provincia_Desc <> "" Then
   Call sbCboAsignaDato(cboTraProvincia, rs!Tra_Provincia_Desc & "")  'Se activa el Click ->    Call cboProvincia_Click
  End If
   
  If rs!Tra_Canton_Desc <> "" Then
   Call sbCboAsignaDato(cboTraCanton, rs!Tra_Canton_Desc & "")  'Se activa el Click ->    Call cboProvincia_Click
  End If
   
  If rs!Tra_Canton_Desc <> "" Then
   Call sbCboAsignaDato(cboTraDistrito, rs!Tra_Canton_Desc & "")  'Se activa el Click ->    Call cboProvincia_Click
  End If
   
  txtTraDireccion.Text = rs!Tra_Direccion & ""
   
   txtEmail = Trim(rs!AF_Email & "")
   txtEmail_02 = Trim(rs!Email_02 & "")
   txtApartado = Trim(rs!apto & "")
   
'   If IIf(IsNull(rs!EstadoLaboral), 1, rs!EstadoLaboral) = 1 Then
'     optNombramiento.Item(0).Value = True
'   Else
'     optNombramiento.Item(1).Value = True
'   End If
'
   Call sbCboAsignaDato(cboEstadoLaboral, rs!EstadoLaboralDesc, True, rs!estadoLaboral)
   
   dtpNombramiento.Value = IIf(IsNull(rs!nombramiento_fecha), dtpFechaIngreso.Value, rs!nombramiento_fecha)
   
   txtAniosSerivicio.Text = Trim(rs!AnioServicio)
       
   'Ultimo Salario
   txtSalarioDevengado.Text = Format(rs!SALARIO_DEVENGADO, "Standard")
   dtpSalarioFecha.Value = IIf(IsNull(rs!Salario_fecha), dtpFechaIngreso.Value, rs!Salario_fecha)
   txtSalarioNeto.Text = Format(rs!Salario_Neto, "Standard")
   txtSalarioRebajos.Text = Format(rs!SALARIO_REBAJOS, "Standard")
       
   chkSalarioEmbargos.Value = rs!SalarioEmbargo
    
   If Not IsNull(rs!TIPO_SALARIO) Then
        Call sbCboAsignaDato(cboSalarioTipo, rs!Salario_Tipo, True, rs!TIPO_SALARIO)
   End If
       
   If Not IsNull(rs!Salario_Divisa) Then
        Call sbCboAsignaDato(cboSalarioDivisa, rs!SalarioDivisaDesc, True, rs!Salario_Divisa)
   End If
       
   txtConyugeCedula.Text = Trim(rs!conyuge_cedula & "")
   txtConyugeNombre.Text = Trim(rs!conyuge_nombre & "")
   txtConyugeTelCelular.Text = Trim(rs!conyuge_TelCell & "")
   txtConyugeTelTrabajo.Text = Trim(rs!conyuge_TelTra & "")
   txtConyugeTelTrabajoExt.Text = Trim(rs!conyuge_TelTraExt & "")
   
   txtAlbaceaCedula.Text = Trim(rs!albacea_Cedula & "")
   txtAlbaceaNombre.Text = Trim(rs!albacea_nombre & "")
   txtAlbaceaTelCelular.Text = Trim(rs!albacea_TelCell & "")
   txtAlbaceaTelTrabajo.Text = Trim(rs!ALBACEA_TELTRA & "")
   txtAlbaceaTelTrabajoExt.Text = Trim(rs!albacea_TelTraExt & "")
   
   
   
   txtPromotorCod.Text = rs!ID_PROMOTOR
   txtPromotorDesc.Text = Trim(rs!promotor)
   
   txtNotificaciones.Text = Trim(rs!Notificaciones & "")
   
   
   txtProfesionCod.Text = rs!Cod_Profesion
   txtProfesionDesc.Text = rs!ProfesionDesc
   
   txtSectorCod.Text = rs!Cod_Sector
   txtSectorDesc.Text = rs!SectorDesc
   
   
   txtInstitucionCod.Text = rs!cod_institucion
   txtInstitucionDesc.Text = rs!InstitucionDesc
   
   txtDeduccionesCod.Text = Trim(rs!DeductoraCod)
   txtDeduccionesDesc.Text = rs!DeductoraDesc
   
   
    txtDeptCodigo.Text = Trim(rs!UP & "")
    txtDeptDesc.Text = Trim(rs!UP_esc & "")
    
    txtSecCodigo.Text = Trim(rs!UT & "")
    txtSecDesc.Text = Trim(rs!UT_Desc & "")
   
   If Not GLOBALES.SysASEVersion Then
        txtDeptCodigo.Text = Trim(rs!cod_departamento & "")
        txtDeptDesc.Text = Trim(rs!DepartamentoDesc & "")
        
        txtSecCodigo.Text = Trim(rs!cod_seccion & "")
        txtSecDesc.Text = Trim(rs!SeccionDesc & "")
   End If
   
   txtCT.Text = Trim(rs!CT & "")
   txtCT_Desc.Text = Trim(rs!CT_Desc & "")
   
   If Not IsNull(rs!Nivel_Academico) Then
    Call sbCboAsignaDato(cboNivelAcademico, rs!NivelAcademicoDesc, False, rs!Nivel_Academico)
   End If
       
   txtPuestoDesc.Text = Trim(rs!COD_CARGO & "")
   
   If Not IsNull(rs!I_TRABAJO_PROPIO) Then
    chkTrabajoPropio.Value = rs!I_TRABAJO_PROPIO
   End If
   
   If Not IsNull(rs!I_BENEFICIARIOS) Then
    chkBeneficiarios.Value = rs!I_BENEFICIARIOS
   End If
   
   If Not IsNull(rs!TIPO_PATRON) Then
     Select Case rs!TIPO_PATRON
        Case 0 'Publico
            cboPatrono.Text = "Público"
        Case 1 'Privado
            cboPatrono.Text = "Privado"
        Case 2 'Otro
            cboPatrono.Text = "Otro"
     End Select
   End If
   
   If Not IsNull(rs!Tipo_CES) Then
      If rs!Tipo_CES < 3 Then
        rbCES(rs!Tipo_CES - 1).Value = True
      End If
   End If
   
   vPaso = True
   
   If Not IsNull(rs!PEP_Ind) Then
        chkPersonaPolitica.Value = rs!PEP_Ind
        dtpC_CargoInicio.Value = IIf(IsNull(rs!PEP_Inicio), vFechaActual, rs!PEP_Inicio)
        dtpC_CargoCorte.Value = IIf(IsNull(rs!PEP_Corte), vFechaActual, rs!PEP_Corte)
        txtC_CargoPolitico.Text = rs!PEP_Cargo & ""
   End If
   
   
   txtHijos.Text = IIf(IsNull(rs!hijos), 0, rs!hijos)
   txtNumeroPagos = IIf(IsNull(rs!af_npagos), 0, rs!af_npagos)
   lblOficina.Caption = IIf(IsNull(rs!Oficina), "Sin Descripción", rs!Oficina)
   
   If lblOficina.Caption = "Sin Descripción" Then
    lblOficina.Tag = 0
   Else
    lblOficina.Tag = rs!COD_OFICINA
   End If
   
   
       chkBienes.Value = IIf(IsNull(rs!ind_propiedades), 0, rs!ind_propiedades)
       
       chkBloqueo.Value = IIf(IsNull(rs!bloqueo), 0, rs!bloqueo)
       chkDesactivaAporte.Value = IIf(IsNull(rs!ind_sinAporte), 0, rs!ind_sinAporte)
       chkDobleDeduccion.Value = IIf(IsNull(rs!IND_DOBLE_DEDUCCION), 0, rs!IND_DOBLE_DEDUCCION)
    
       If IsNull(rs!Consentimiento_Contacto_Fecha) Then
           chkConsentimiento.Value = xtpUnchecked
       Else
           chkConsentimiento.Value = xtpChecked
       End If
       
       chkAportePatronalAdministra.Value = IIf(IsNull(rs!AutorizaAdminAportePatronal), 0, rs!AutorizaAdminAportePatronal)
   
   vPaso = False
   
   txtNotasAdv.Text = rs!Notas & ""
   
   chkDimex_Activo.Value = IIf(IsNull(rs!Dimex_Activo), 0, rs!Dimex_Activo)
   txtDIMEX.Text = Trim(rs!Dimex_Cedula & "")
   txtDimex_Nuevo.Text = Trim(rs!Dimex_Cedula & "")
    
   txtDimex_RUsuario.Text = Trim(rs!Dimex_Usuario & "")
   txtDimex_RFecha.Text = Trim(rs!Dimex_Fecha & "")
   
   txtDimex_AUsuario.Text = Trim(rs!Dimex_Actualiza_Usuario & "")
   txtDimex_AFecha.Text = Trim(rs!Dimex_Actualiza_Fecha & "")
    
   
   
   
   txtSN_Blog.Text = Trim(rs!Blog & "")
   txtSN_Facebook.Text = Trim(rs!Facebook & "")
   txtSN_Twitter.Text = Trim(rs!TWITTER & "")
   txtSN_LinkedIn.Text = Trim(rs!Linkedin & "")
   txtSN_Instagram.Text = Trim(rs!Instagram & "")
   
   txtCedula.SetFocus
   
   StatusBarX.Panels.Item(1) = rs!reg_user & ""
   StatusBarX.Panels.Item(2) = rs!reg_fecha & ""
   StatusBarX.Panels.Item(3) = rs!ActualizaUser & ""
   StatusBarX.Panels.Item(4) = rs!ActualizaFecha & ""
   
   vCambios.vFecNac = dtpNacimiento.Value ' carga fecha Nac. para verificacion
   vCambios.vEstado = cboEstado.Text ' carga Estado civil para verificacion
   vCambios.vPromotor = txtPromotorDesc.Text ' carga promotor para verificacion

 
'   gbOpciones.Visible = True
 
  Else
   
   If vEditar Then
        vEditar = False
        Call sbBarra_Accion("nuevo")
        Call RefrescaTags(Me)
        Call sbClearControles
        Call sbLockControles("L")
        txtCedula.SetFocus
   Else
        Call RefrescaTags(Me)
        
        Call sbLimpiaDatos
        
        If fraTipo.Visible Then
           txtNombreComercial.SetFocus
        Else
           txtApellido1.SetFocus
        End If
        
        
        'Consulta Padron
        Call gBase_Padron(txtCedula.Text, "General", rs, "CRI")
        
        If rs.RecordCount > 0 Then
           txtApellido1.Text = Trim(rs!Apellido_1)
           txtApellido2.Text = Trim(rs!Apellido_2)
           txtNombre.Text = Trim(rs!Nombre)
           cboSexo.Text = IIf(rs!sexo = "F", "Femenino", "Masculino")
           
            txtDireccion = Trim(rs!direccion) & ""
            txtEmail.Text = Trim(rs!Email_01 & "")
            txtEmail_02.Text = Trim(rs!Email_02 & "")
            
            dtpNacimiento.Value = rs!fecha_nacimiento
            cboEstado.Text = fxEstadoCivil(rs!Estado_Civil & "")
            
            Call sbCboAsignaDato(cboProvincia, rs!Provincia & "", False)  'Se activa el Click ->    Call cboProvincia_Click
            Call sbCboAsignaDato(cboCanton, rs!Canton & "", False)   'Se activa el Click
            Call sbCboAsignaDato(cboDistrito, rs!Distrito & "", False)
        Else
            
            MsgBox "No se encontró la persona en el Padron, continue con el registro manual!", vbExclamation

        End If
        
        
   End If

End If
rs.Close

btnEditarDetalle.Enabled = True

End Sub


Sub sbLockControles(vModo As String)
'Dim vControl As Control
'
'For Each vControl In Me
'  If (TypeOf vControl Is TextBox And vControl.Name <> "txtCedula" And vControl.Name <> _
'     "txtEstado") Or TypeOf vControl Is DTPicker Or TypeOf vControl Is ComboBox Then
'        If vModo = "L" Then
'           If vControl.Name = "txtNombre" Then
'            vControl.Locked = True
'           Else
'            vControl.Enabled = False
'           End If
'        Else
'           If vControl.Name = "txtNombre" Then
'            vControl.Locked = False
'           Else
'            vControl.Enabled = True
'           End If
'        End If
'  End If
'Next

dtpFechaIngreso.Enabled = False
cboTipoId.Enabled = True

End Sub


Private Sub sbDeleteRecord()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If Trim(txtCedula) <> vCedula Then
   MsgBox "Ha modificado la cédula", vbExclamation
   Exit Sub
End If

strSQL = "exec spAFIPersonaMovAux '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Total > 0 Then
 strSQL = "Tiene Referencias en Otros Módulos..:" & vbCrLf & vbCrLf _
        & "Patrimonio...:" & rs!Patrimonio & vbCrLf _
        & "Fondos    ...:" & rs!fondos & vbCrLf _
        & "Créditos  ...:" & rs!Creditos & vbCrLf _
        & "Fianzas   ...:" & rs!FIANZAS
        
  MsgBox strSQL, vbExclamation
  rs.Close
  Exit Sub
        
 End If
rs.Close

i = MsgBox("Esta Seguro Que Desea Borrar esta Persona?", vbYesNo)
If i = vbYes Then
  vEditar = False
  strSQL = "delete Socios where Cedula='" & Trim(txtCedula) & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Borra", "Persona [Identificación : " & Trim(txtCedula) & "]")
  
  Call sbClearControles
  Call sbBarra_Accion("nuevo")
  Call RefrescaTags(Me)
  txtCedula.SetFocus
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

'Funcion para Apellidos y Nombres que no contenga números ni otro tipo de caracter que no sean letras del alfabeto
Private Function fxValidaCadena(pCadena As String) As Boolean
Dim str As String, i As Integer, vResultado As Boolean

vResultado = True
pCadena = Trim(pCadena)
For i = 1 To Len(pCadena)
  vResultado = False
  str = Mid(pCadena, i, 1)
   If Asc(str) >= 65 And Asc(str) <= 90 Then
     vResultado = True
   End If
  
   If Asc(str) >= 97 And Asc(str) <= 122 Then
     vResultado = True
   End If
  
   If str = "á" Or str = "é" Or str = "í" Or str = "ó" Or str = "ú" Then
     vResultado = True
   End If
   
   If str = "Á" Or str = "É" Or str = "Í" Or str = "Ó" Or str = "Ú" Then
     vResultado = True
   End If
   
   If str = " " Or str = "ñ" Or str = "Ñ" Then
     vResultado = True
   End If
   
   If Not vResultado Then Exit For
Next i

fxValidaCadena = vResultado

End Function

Private Function fxValida() As Boolean
Dim rs As New ADODB.Recordset, i As Integer, pFecha As Date
Dim vMensaje As String

vMensaje = ""
 
 pFecha = fxFechaServidor
 
'Actualiza el Parametro de Validacion y Luego lo Aplica
strSQL = "select LARGO_MINIMO from AFI_TIPOS_IDS Where TIPO_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    vParametros.LargoCedula = rs!Largo_Minimo
End If
rs.Close

If Not vEditar Then
    If Len(Trim(txtCedula)) <> vParametros.LargoCedula Then vMensaje = vMensaje & " - Número de Identidad no es válido, se espera que sea de: " & vParametros.LargoCedula _
            & " caracteres, verifique...!" & vbCrLf
End If

If Len(Trim(txtCedula)) > 20 Then vMensaje = vMensaje & " - Número de Identidad no es válido, verifique...!" & vbCrLf

If Not fxEmail_Valida(txtEmail.Text) Then
    vMensaje = vMensaje & " - El Email principal no es válido!" & vbCrLf
End If

If Len(Trim(txtEmail_02.Text)) > 0 Then
    If Not fxEmail_Valida(txtEmail_02.Text) Then
        vMensaje = vMensaje & " - El Email secundario no es válido!" & vbCrLf
    End If
End If

'validacion segun tipo de personeria
If fraTipo.Visible Then
 'personeria juridica
 If Trim(txtNombreComercial.Text) = "" Then vMensaje = vMensaje & " - Indique el nombre comercial" & vbCrLf
 If Trim(txtRazonSocial.Text) = "" Then vMensaje = vMensaje & " - Indique la Razón Social de la Sociedad" & vbCrLf

Else
 'persona fisica
    If Trim(txtApellido1) = "" Then vMensaje = vMensaje & " - Falta el Apellido 1" & vbCrLf
    If Trim(txtApellido2) = "" Then vMensaje = vMensaje & " - Falta el Apellido 2" & vbCrLf
    If Trim(txtNombre) = "" Then vMensaje = vMensaje & " - Falta el Nombre" & vbCrLf
    
    
    If Not fxValidaCadena(txtApellido1.Text) Then vMensaje = vMensaje & " - El Apellido 1: No es válido" & vbCrLf
    If Not fxValidaCadena(txtApellido2.Text) Then vMensaje = vMensaje & " - El Apellido 2: No es válido" & vbCrLf
    If Not fxValidaCadena(txtNombre.Text) Then vMensaje = vMensaje & " - El Nombre: No es válido" & vbCrLf

    If Trim(cboSexo) = "" Then vMensaje = vMensaje & " - No se especificó el Sexo" & vbCrLf
    If Trim(cboEstado) = "" Then vMensaje = vMensaje & " - No se especificó el Estado Civil" & vbCrLf
End If

If cboSalarioDivisa.ItemData(cboSalarioDivisa.ListIndex) = "COL" Then
    If CCur(txtSalarioDevengado.Text) < 100000 Or CCur(txtSalarioDevengado.Text) > 10000000 Then vMensaje = vMensaje & " - Salario Devengado no es válido" & vbCrLf
Else
    If CCur(txtSalarioDevengado.Text) < 200 Or CCur(txtSalarioDevengado.Text) > 20000 Then vMensaje = vMensaje & " - Salario Devengado no es válido" & vbCrLf
End If

If Trim(cboProvincia.Text) = "" Then vMensaje = vMensaje & " - No se especificó la Provincia" & vbCrLf
If Trim(cboCanton.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Cantón" & vbCrLf
If Trim(cboDistrito.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Distrito en la dirección" & vbCrLf

'If Trim(txtDireccion) = "" Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf
If Not fxDireccion_Valida(Trim(txtDireccion), "-,#,*") Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf

'Si se indica una direccion de Trabajo esta será revisada
If Len(txtTraDireccion.Text) > 0 Then
    If Trim(cboTraProvincia.Text) = "" Then vMensaje = vMensaje & " - No se especificó la Provincia de la Dirección de Trabajo" & vbCrLf
    If Trim(cboTraCanton.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Cantón de la Dirección de Trabajo" & vbCrLf
    If Trim(cboTraDistrito.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Distrito en la Dirección de Trabajo" & vbCrLf
    
    If Not fxDireccion_Valida(Trim(txtTraDireccion), "-,#,*") Then vMensaje = vMensaje & " - La Dirección Exacta de Trabajo no fue suministrada correctamente!" & vbCrLf

End If

If Trim(txtHijos) = "" Then txtHijos.Text = 0
If Trim(txtBoleta) = "" Then txtBoleta.Text = 0
If Trim(txtNumeroPagos) = "" Then txtNumeroPagos.Text = 0

If Not IsNumeric(txtBoleta.Text) Then vMensaje = vMensaje & " - Número de Boleta no es válido" & vbCrLf
If Not IsNumeric(txtHijos.Text) Then vMensaje = vMensaje & " - Número de Hijos no es válido" & vbCrLf
If Not IsNumeric(txtNumeroPagos.Text) Then vMensaje = vMensaje & " - Número de Pagos no es válido" & vbCrLf


If Not IsNumeric(txtProfesionCod.Text) Then vMensaje = vMensaje & " - Profesion no es válida!" & vbCrLf
If Not IsNumeric(txtSectorCod.Text) Then vMensaje = vMensaje & " - Sector no es válido" & vbCrLf

If Trim(cboEstadoLaboral.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Estado Laboral" & vbCrLf
If Trim(cboNivelAcademico.Text) = "" Then vMensaje = vMensaje & " - No se especificó el nivel academico" & vbCrLf
If Trim(cboActividad.Text) = "" Then vMensaje = vMensaje & " - No se especificó Actividad (Oficina Cumplimiento)" & vbCrLf

If lblDepartamento.Caption = "U.Programatica" Then
    If Trim(txtDeptCodigo.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Departamento o Unidad Programatica" & vbCrLf
End If

If Len(txtPuestoDesc.Text) = 0 Then vMensaje = vMensaje & " - Tienen que indicar el Puesto que desempeña" & vbCrLf
'Verificar el Estado de la Persona si está autorizado en la institución
strSQL = "select count(*) as Existe from AFI_ESTADOS_INSTITUCIONES" _
       & " where cod_estado = '" & cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) & "' and cod_institucion = " & txtDeduccionesCod.Text
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = vMensaje & " - El ESTADO de la Persona a modificar o incluir no está autorizado en esta institución: " & txtInstitucionCod.Text & vbCrLf
End If
rs.Close

If Not vEditar Then
  strSQL = "select estado from promotores where id_promotor = " & txtPromotorCod & ""
  Call OpenRecordSet(rs, strSQL)
  If rs!estado = 0 Then
    vMensaje = vMensaje & " - El promotor indicado se encuenta inactivo" & vbCrLf
  End If
  rs.Close

  If vTipoJuridica <> cboTipoId.ItemData(cboTipoId.ListIndex) Then
      If dtpNacimiento.Value > DateAdd("yyyy", -17, pFecha) Then
        vMensaje = vMensaje & " - Verifique la Fecha de nacimiento, la persona es menor de edad!" & vbCrLf
      End If
    
      If dtpCedulaVence.Value <= DateAdd("d", 20, pFecha) Then
        vMensaje = vMensaje & " - Verifique la fecha de Vencimiento del documento de Identidad, está pronta a vencer!" & vbCrLf
      End If
  End If
  
'- Verifica que no existe otra persona con el mismo nombre -> Solo Adventencia
    'Filtra nombre
    
    If vParametros.VerificaNombre Then
        strSQL = ""
        For i = 1 To 4 'Len(txtNombre)
          If Mid(txtNombre, i, 1) <> " " Then
             strSQL = strSQL & Mid(txtNombre, i, 1)
          Else
             Exit For
          End If
        Next i
      
        strSQL = "select isnull(count(*),0) as Existe from socios where nombre like '" & Trim(txtApellido1.Text) _
               & " " & Trim(txtApellido2.Text) & " " & strSQL & "%'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe > 0 Then
          i = MsgBox("Existe otra persona con el mismo nombre, está seguro de que sea continuar con el registro?", vbYesNo)
          If i = vbNo Then vMensaje = vMensaje & " - Existe otra persona con el mismo nombre, verifique..." & vbCrLf
        End If
        rs.Close
    End If

   If vParametros.VerificaPadron Then
      strSQL = "select isnull(count(*),0) as Existe from afi_padron where cedula = '" & Trim(txtCedula.Text) _
             & "' and institucion = " & txtInstitucionCod.Text
      Call OpenRecordSet(rs, strSQL)
      If rs!Existe = 0 Then
        i = MsgBox("No Existe esta persona registrada como Empleado en :" & txtInstitucionCod.Text & ", desea continuar...!", vbYesNo)
        If i = vbNo Then vMensaje = vMensaje & " - No existe registro en el nomina de la institucion correspondiente a persona, verifique..." & vbCrLf
      End If
      rs.Close
   End If

End If 'Insertar


strSQL = "select dbo.fxAFI_Afiliacion_Valida_Beneficiarios('" & txtCedula.Text & "') as 'Beneficiarios'" _
       & ", dbo.fxAFI_Afiliacion_Valida_Telefonos('" & txtCedula.Text & "') as 'Telefonos'" _
       & ", dbo.fxAFI_Afiliacion_Valida_Beneficiarios_MenoresSinAlbacea('" & txtCedula.Text & "') as 'MenoresSinAlbacea'"
Call OpenRecordSet(rs, strSQL)
If rs!Beneficiarios = 0 And chkBeneficiarios.Value = xtpChecked Then
      vMensaje = vMensaje & " - No se han registrado los Beneficiarios o Estan incompletos!" & vbCrLf
End If

If rs!MenoresSinAlbacea > 0 And chkBeneficiarios.Value = xtpChecked And Len(txtAlbaceaCedula.Text) <= 5 Then
      vMensaje = vMensaje & " - Existen Beneficiarios Menores de Edad y no se han indicado lo(s) Albacea(s)!" & vbCrLf
End If

If rs!Telefonos = 0 Then
      vMensaje = vMensaje & " - No se han registrado los telefonos de contacto!" & vbCrLf
End If

If Len(vMensaje) = 0 Then
  fxValida = True
Else
  fxValida = False
  MsgBox vMensaje, vbExclamation

End If

End Function

Private Function fxVerificaCambioInst(xCedula As String, vInst As Integer) As Boolean
Dim vSQL As String, rsX As New ADODB.Recordset

fxVerificaCambioInst = True

vSQL = "select A.aporte,S.cod_institucion" _
     & " from ahorro_consolidado A inner join Socios S on A.cedula = S.cedula" _
     & " where S.cedula = '" & xCedula & "'"
rsX.Open vSQL, glogon.Conection, adOpenStatic
If Not rsX.EOF And Not rsX.BOF Then
   If rsX!cod_institucion <> vInst Then
     If rsX!Aporte > 0 Then fxVerificaCambioInst = False
     If rsX!Aporte = 0 Then
            fxVerificaCambioInst = True
            If vParametros.BitacoraEspecial Then
               Call sbgAFIBitacora("01", "Cambio de Institucion para Persona:  " & txtCedula.Text, Trim(txtCedula.Text))
            End If
     End If
   End If
End If
rsX.Close

End Function


Private Sub sbSaveRecord()
Dim vEstadoCivil As String, vActiva As Boolean
Dim i As Integer, vNombre As String

Dim pTipoCES As Integer
Dim vUp As String, vCt As String, vUT As String
Dim vActividades As String

Dim vMov As String

On Error GoTo vError

strSQL = ""
vActiva = False
mSavePass = False

If Not fxValida Then
  Exit Sub
End If

If vEditar = True Then
  
 If Trim(txtCedula) <> vCedula Then
   MsgBox "Ha modificado la identificación", vbExclamation
   Exit Sub
 End If
 'Verifica que pueda cambiar de institucion
 If fxVerificaCambioInst(txtCedula, txtInstitucionCod.Text) = False Then
   MsgBox "No se puede cambiar la institución a esta persona porque ya tiene aportes registrados: debe liquidar primero", vbExclamation
   Exit Sub
 End If
 
End If

If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
    vEstadoCivil = "O"
Else
    vEstadoCivil = cboEstado.ItemData(cboEstado.ListIndex)
End If


'Nombre comercial o de persona fisica
If fraTipo.Visible Then
   vNombre = txtNombreComercial.Text
Else
    strSQL = ""
    For i = 1 To Len(txtApellido1)
      If Mid(txtApellido1, i, 1) <> " " Then
         strSQL = strSQL & Mid(txtApellido1, i, 1)
      End If
    Next i
    txtApellido1 = strSQL
    
    strSQL = ""
    For i = 1 To Len(txtApellido2)
      If Mid(txtApellido2, i, 1) <> " " Then
         strSQL = strSQL & Mid(txtApellido2, i, 1)
      End If
    Next i
    txtApellido2 = strSQL
   
   vNombre = Trim(txtApellido1) & " " & Trim(txtApellido2) & " " & Trim(txtNombre)
End If

'Validacion de Numericos
If Not IsNumeric(txtSalarioDevengado.Text) Then txtSalarioDevengado.Text = "0"
If Not IsNumeric(txtSalarioNeto.Text) Then txtSalarioNeto.Text = "0"
If Not IsNumeric(txtSalarioRebajos.Text) Then txtSalarioRebajos.Text = "0"



If GLOBALES.SysASEVersion Then
   vUp = txtDeptCodigo.Text
   vUT = txtSecCodigo.Text
   vCt = txtCT.Text
Else
   vUp = ""
   vUT = ""
   vCt = ""
End If

If vEditar Then
  vMov = "E"
Else
  vMov = "A"
    txtPrideduc.Text = fxgPrimerDeduccionIng(txtDeduccionesCod.Text)
End If


'@TipoId int, @Cedula varchar(20), @Id_Alterno varchar(20), @Nombre_Completo varchar(150), @Apellido_1 varchar(30), @Apellido_2 varchar(30), @Nombre varchar(30)
'                    , @RazonSocial varchar(150), @Estado varchar(10), @EstadoCivil varchar(10), @Genero varchar(10), @fNacimiento datetime, @fCedulaVence datetime
'                    , @PromotorId int, @Boleta varchar(30), @fIngreso datetime, @EstadoLaboral varchar(10)
'                    , @PaisNac varchar(10), @Nacionalidad varchar(10), @Email_1 varchar(150), @Email_2 varchar(150)
'                    , @Provincia varchar(10), @Canton varchar(10), @Distrito varchar(20), @Direccion varchar(500), @AptoPostal varchar(30), @Notificacion varchar(500)
'                    , @Institucion int, @Departamento varchar(10), @Seccion varchar(10), @UP varchar(10), @UT varchar(10), @CT varchar(10), @Deductora int
'                    , @Profesion  int, @Sector int, @NPagos smallint, @NHijos smallint, @PriDeduc dec(7,1), @fNombramiento datetime, @NivelAcademico int
'
'                    , @Sociedad varchar(10), @Actividad varchar(10), @Propiedades smallint, @Oficina varchar(10)
'                    , @facebook varchar(200), @Twitter varchar(200), @LinkedIn varchar(200), @Instagram varchar(200), @Blog varchar(200)
'
'                    , @ConyugeCedula varchar(20), @ConyugeNombre varchar(100), @ConyugeTelCel varchar(15), @ConyugeTelTra varchar(15), @ConyugeTelTraExt varchar(15)
'                    , @AlbaceaCedula varchar(20), @AlbaceaNombre varchar(100), @AlbaceaTelCel varchar(15), @AlbaceaTelTra varchar(15), @AlbaceaTelTraExt varchar(15)
'
'                    , @SalarioTipo varchar(10), @SalarioDivisa varchar(10), @SalarioFecha datetime, @SalarioDevengado dec(16,2), @SalarioRebajos dec(16,2), @SalarioNeto dec(16,2), @SalarioEmbargo char(1)
'
'                    , @AdminitraAportePatronal smallint, @Sugef smallint, @I_Beneficiario bit, @I_TrabajoPropio bit, @Tipo_Patron varchar(10),   @CargoDesc varchar(200)
'                    , @PEP_Ind smallint, @PEP_Inicio datetime, @PEP_Corte datetime, @PEP_Cargo varchar(200), @TipoCES smallint
'                    , @Usuario varchar(30), @Mov char(1) = 'A')


Select Case True
    Case rbCES(0).Value
        pTipoCES = 1
    Case rbCES(1).Value
        pTipoCES = 2
    Case rbCES(2).Value
        pTipoCES = 3
End Select

If cboActividad.Text <> "" Then
    vActividades = cboActividad.ItemData(cboActividad.ListIndex)
Else
    vActividades = "Null"
End If


Dim pTraProvincia As String, pTraCanton As String, pTraDistrito As String, pTraDireccion As String

txtTraDireccion.Text = fxSysCleanTxtInject(txtTraDireccion.Text)

If Len(txtTraDireccion.Text) = 0 Then
   pTraProvincia = "Null"
   pTraCanton = "Null"
   pTraDistrito = "Null"
   pTraDireccion = "Null"
Else
   pTraProvincia = "'" & cboTraProvincia.ItemData(cboTraProvincia.ListIndex) & "'"
   pTraCanton = "'" & cboTraCanton.ItemData(cboTraCanton.ListIndex) & "'"
   pTraDistrito = "'" & cboTraDistrito.ItemData(cboTraDistrito.ListIndex) & "'"
   pTraDireccion = "'" & Mid(txtTraDireccion.Text, 1, 1000) & "'"
End If


strSQL = "exec spAFI_Persona_Add " & cboTipoId.ItemData(cboTipoId.ListIndex) & ",'" & Trim(txtCedula) & "', '" & Trim(txtCedAlternativa.Text) & "', '" & vNombre & "', '" & Trim(txtApellido1.Text) _
       & "', '" & Trim(txtApellido2.Text) & "', '" & Trim(txtNombre.Text) & "', '" & Trim(txtRazonSocial.Text) & "', '" & cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) _
       & "', '" & vEstadoCivil & "', '" & Mid(cboSexo.Text, 1, 1) & "', '" & Format(dtpNacimiento.Value, "yyyy/mm/dd") & "', '" & Format(dtpCedulaVence.Value, "yyyy/mm/dd") _
       & "',  " & txtPromotorCod.Text & ", '" & txtBoleta.Text & "', '" & Format(dtpFechaIngreso.Value, "yyyy/mm/dd") & "', '" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) _
       & "', '" & cboPaisNac.ItemData(cboPaisNac.ListIndex) & "', '" & cboNacionalidad.ItemData(cboNacionalidad.ListIndex) & "', '" & Trim(txtEmail.Text) & "', '" & Trim(txtEmail_02.Text) _
       & "', '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "', '" & cboCanton.ItemData(cboCanton.ListIndex) & "', '" & cboDistrito.ItemData(cboDistrito.ListIndex) _
       & "', '" & Trim(txtDireccion.Text) & "', '" & Trim(txtApartado.Text) & "', '" & Trim(txtNotificaciones.Text) _
       & "',  " & txtInstitucionCod.Text & ", '" & txtDeptCodigo.Text & "', '" & txtSecCodigo.Text & "', '" & vUp & "', '" & vUT & "', '" & vCt _
       & "',  " & txtDeduccionesCod.Text & ", " & txtProfesionCod.Text & ", " & txtSectorCod.Text & ", " & txtNumeroPagos.Text & ", " & txtHijos.Text & ", " & txtPrideduc.Text _
       & ",  '" & Format(dtpNombramiento.Value, "yyyy/mm/dd") & "', " & cboNivelAcademico.ItemData(cboNivelAcademico.ListIndex) _
       & ",  '" & txtTipoSociedadCod.Text & "', '" & txtActividadCod.Text & "', " & chkBienes.Value & ", '" & GLOBALES.gOficinaTitular _
       & "', '" & txtSN_Facebook.Text & "', '" & txtSN_Twitter.Text & "', '" & txtSN_LinkedIn.Text & "', '" & txtSN_Instagram.Text & "', '" & txtSN_Blog.Text _
       & "', '" & txtConyugeCedula.Text & "', '" & txtConyugeNombre.Text & "', '" & txtConyugeTelCelular.Text & "', '" & txtConyugeTelTrabajo.Text & "','" & txtConyugeTelTrabajoExt.Text _
       & "', '" & txtAlbaceaCedula.Text & "', '" & txtAlbaceaNombre.Text & "', '" & txtAlbaceaTelCelular.Text & "', '" & txtAlbaceaTelTrabajo.Text & "','" & txtAlbaceaTelTrabajoExt.Text _
       & "', '" & cboSalarioTipo.ItemData(cboSalarioTipo.ListIndex) & "', '" & cboSalarioDivisa.ItemData(cboSalarioDivisa.ListIndex) & "', '" & Format(dtpSalarioFecha.Value, "yyyy/mm/dd") _
       & "',  " & CCur(txtSalarioDevengado.Text) & ", " & CCur(txtSalarioRebajos.Text) & ", " & CCur(txtSalarioNeto.Text) & ",  '" & IIf(chkSalarioEmbargos.Value = xtpChecked, "S", "N") _
       & "',  " & chkAportePatronalAdministra.Value & ", 0, " & chkBeneficiarios.Value & ", " & chkTrabajoPropio.Value & ", '" & cboPatrono.ItemData(cboPatrono.ListIndex) _
       & "', '" & txtPuestoDesc.Text & "', " & chkPersonaPolitica.Value & ", '" & Format(dtpC_CargoInicio.Value, "yyyy-mm-dd") & "', '" & Format(dtpC_CargoCorte.Value, "yyyy-mm-dd") _
       & "', '" & txtC_CargoPolitico.Text & "', " & pTipoCES & ", " & vActividades & ", '" & glogon.Usuario & "', '" & vMov & "'" _
       & ", " & pTraProvincia & ", " & pTraCanton & ", " & pTraDistrito & ", " & pTraDireccion
Call OpenRecordSet(rs, strSQL)
   
If rs!Pass = 0 Then
   Me.MousePointer = vbDefault
   MsgBox rs!Error_Msj, vbExclamation
   Exit Sub
End If
vCedula = Trim(txtCedula.Text)

Call Bitacora(IIf(vMov = "A", "Registra", "Modifica"), "Número Identificación: " & vCedula)
If vEditar Then
  vActiva = True
Else
  vActiva = False
End If

If Not vEditar Then
        'Boleta de Afiliacion
        Call btnInformes_Click(0)
        'Contrato Cuenta Sinpe
        Call sbFnd_Contratos_Cuenta_Sinpe(txtCedula.Text)
End If

txtDeptCodigo.Text = ""
txtSecCodigo.Text = ""

Call sbCurrentRecord


Call sbBarra_Accion("activo")
Call RefrescaTags(Me)
Call sbLockControles("L")

Me.MousePointer = vbDefault
MsgBox "Información guardada satisfactoriamente...", vbInformation
txtCedula.SetFocus

'Abre el Marco de Contacto
Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)

GLOBALES.gCedulaActual = Trim(txtCedula)
If vParametros.SolicitarTelefonos And Not vEditar Then
  frmAF_Telefonos.Show vbModal
  
End If

If chkBeneficiarios.Value = xtpChecked Then
    If vParametros.SoliciarBeneficiario And Not vEditar Then
      frmAF_Beneficiarios.Show vbModal
    End If
End If

If vParametros.SolicitarCuentas And Not vEditar Then
    GLOBALES.gTag = Trim(txtCedula.Text)
    GLOBALES.gTag2 = "AFI"
    frmCC_Cuentas_Bancarias.Show vbModal
End If

vEditar = True

'gbOpciones.Visible = True
mSavePass = True

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbDatosContatoHistorico_Clean()
txtDir_Id.Text = "0"

txtDir_Email1.Text = ""
txtDir_Email2.Text = ""

txtDir_Provincia.Text = ""
txtDir_Canton.Text = ""
txtDir_Distrito.Text = ""
txtDir_Direccion.Text = ""

txtDir_Telefono1.Text = ""
txtDir_Telefono2.Text = ""

txtDir_Usuario.Text = ""
txtDir_Fecha.Text = ""
End Sub


Private Sub btnAdjuntos_Click()

If txtCedula.Text <> "" Then
 gGA.Modulo = "CL_01"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End If

End Sub





Private Sub btnBarra_Click(Index As Integer)

Select Case Index
  Case 0 'nuevo"
    vEditar = False
    Call sbBarra_Accion("edicion")
    Call sbClearControles
    Call sbLockControles("U")
    Call sbLimpiaDatos

    
    txtDeduccionesCod.Enabled = True
    txtDeduccionesDesc.Enabled = True
    FlatScrollBarDeduciones.Enabled = True
 
    txtCedula.SetFocus
    
  Case 1 'editar"
    If Trim(txtCedula) <> vCedula Then
     MsgBox "Ha modificado la cédula", vbExclamation
     Exit Sub
    End If
    
    vEditar = True
    vCedula = Trim(txtCedula)
    Call sbBarra_Accion("edicion")
    Call sbLockControles("U")
    
    Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)

    If txtApellido1.Enabled Then txtApellido1.SetFocus
        
  Case 2 '"borrar"
    Call sbDeleteRecord
        
  Case 3 ' "guardar"
    Call sbSaveRecord
    
  Case 4 '"deshacer"
    vEditar = False
    Call sbBarra_Accion("nuevo")
    Call RefrescaTags(Me)
    Call sbClearControles
    Call sbLockControles("L")
    Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)
    
    txtCedula.SetFocus
  Case 5 'Reporte
    Call btnInformes_Click(1)
    
  Case 6 '"consultar"
    Select Case vSeek
      Case 1, 2
       gBusquedas.Resultado = Trim(txtCedula)
       txtCedula = ""
       vCedula = ""
       gBusquedas.Convertir = "N"
       
       If vSeek = 1 Then
        gBusquedas.Columna = "Cedula"
        gBusquedas.Orden = "Cedula"
       Else
        gBusquedas.Columna = "Nombre"
        gBusquedas.Orden = "Nombre"
       End If
       
       gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
   
       frmBusquedas.Show vbModal
   
       txtCedula = Trim(gBusquedas.Resultado)
       txtCedula_LostFocus
       
      Case 3
       If cboProvincia.Text = "" Then Exit Sub
       gBusquedas.Resultado = ""
       gBusquedas.Resultado2 = ""
   
       gBusquedas.Convertir = "N"
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "Descripcion"
       gBusquedas.Consulta = "Select Canton,Descripcion From Cantones"
       gBusquedas.Filtro = "And Provincia =" & cboProvincia.ItemData(cboProvincia.ListIndex)
   
       frmBusquedas.Show vbModal
   
'       txtCodigoCanton = Trim(gBusquedas.Resultado)
'       txtCanton = Trim(gBusquedas.Resultado2)
    End Select
    
End Select

End Sub

Private Sub btnDimex_Click(Index As Integer)

On Error GoTo vError

Select Case Index
    Case 0 'Actualiza
        strSQL = "exec spAFI_Persona_Dimex_Add '" & Trim(txtCedula.Text) & "','" & Trim(txtDimex_Nuevo.Text) & "', " & chkDimex_Activo.Value & ", '" & glogon.Usuario & "'"
        Call OpenRecordSet(rs, strSQL)
        
        If rs!Pass = 1 Then
            chkDimex_Activo.Value = IIf(IsNull(rs!Dimex_Activo), 0, rs!Dimex_Activo)
            txtDIMEX.Text = Trim(rs!Dimex_Cedula & "")
            txtDimex_Nuevo.Text = Trim(rs!Dimex_Cedula & "")
             
            txtDimex_RUsuario.Text = Trim(rs!Dimex_Usuario & "")
            txtDimex_RFecha.Text = Trim(rs!Dimex_Fecha & "")
            
            txtDimex_AUsuario.Text = Trim(rs!Dimex_Actualiza_Usuario & "")
            txtDimex_AFecha.Text = Trim(rs!Dimex_Actualiza_Fecha & "")
            
            MsgBox "Dimex actualizado satisfactoriamente!", vbInformation
            gbDimeX.Visible = False
        Else
            MsgBox "Verifique los datos del Dimex, puede ser que no cambien o que no cumplan con los caracteres requeridos!", vbExclamation
        End If
    
    
    Case 1 'Cierra
        gbDimeX.Visible = False
      
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnDir_Eliminar_Click()

If txtDir_Id.Tag = "0" Then Exit Sub

Dim strSQL As String, pMensaje As String

On Error GoTo vError

pMensaje = "Datos de Contacto> Histórico> Persona Id: " & txtCedula.Text & ", Linea:" & txtDir_Id.Text

strSQL = "exec spAFI_Persona_Direccion_Elimina '" & txtCedula.Text & "'," _
        & txtDir_Id.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Call Bitacora("Elimina", "Persona > " & pMensaje)

MsgBox "Elimina > " & pMensaje, vbInformation

Call sbDirecciones

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnEditarDetalle_Click()
GLOBALES.gCedulaActual = Trim(txtCedula.Text)

Select Case TituloOpciones.Tag
 Case "Telefonos"
    frmAF_Telefonos.Show vbModal
    Call sbTaskPanel_Accion(Id_TaskItem_Telefonos)
 Case "Beneficiarios"
    
    If Len(txtAlbaceaCedula.Text) > 5 Then
        GLOBALES.gTag = "S"
    Else
        GLOBALES.gTag = "N"
    End If
    
    frmAF_Beneficiarios.Show vbModal
    Call sbTaskPanel_Accion(Id_TaskItem_Beneficiarios)
 Case "Cuentas"
    'frmAF_Bancos.Show vbModal
    GLOBALES.gTag = Trim(txtCedula)
    GLOBALES.gTag2 = "AFI"
    frmCC_Cuentas_Bancarias.Show vbModal
    Call sbTaskPanel_Accion(Id_TaskItem_Cuentas)
 Case "Tarjetas"
    frmAF_PersonaTarjetas.Show vbModal
    Call sbTaskPanel_Accion(Id_TaskItem_Tarjetas)
 Case "Direcciones"
    Call sbDirecciones
End Select

End Sub

Private Sub sbDirecciones()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

tcMain.Item(5).Selected = True

lswDireccion.ListItems.Clear

With lswDireccion.ColumnHeaders
    .Clear
    .Add , , "Fecha", 2500
    .Add , , "Provincia", 1800, vbCenter
    .Add , , "Canton", 1800, vbCenter
    .Add , , "Distrito", 1800, vbCenter
    .Add , , "Dirección", 1800
    .Add , , "Email 01", 2800
    .Add , , "Email 02", 2800
    .Add , , "Tel No 1", 1800, vbCenter
    .Add , , "Tel No 2", 1800, vbCenter
    .Add , , "Usuario", 1800, vbCenter
End With

Call sbDatosContatoHistorico_Clean
    
strSQL = "exec spAFI_PERSONA_DIRECCIONES_Consulta '" & Trim(txtCedula) & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   Set itmX = lswDireccion.ListItems.Add(, , rs!Registro_Fecha & "")
       itmX.SubItems(1) = Trim(rs!ProvinciaDesc)
       itmX.SubItems(2) = Trim(rs!CantonDesc)
       itmX.SubItems(3) = Trim(rs!DistritoDesc)
       itmX.SubItems(4) = Trim(rs!direccion)
       itmX.SubItems(5) = Trim(rs!Email_01)
       itmX.SubItems(6) = Trim(rs!Email_02)
       itmX.SubItems(7) = Trim(rs!Telefono_01)
       itmX.SubItems(8) = Trim(rs!Telefono_02)
       itmX.SubItems(9) = Trim(rs!Registro_Usuario & "")
       itmX.Tag = CStr(rs!Linea_Id)
   rs.MoveNext
Loop
rs.Close
  
  
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExport_Click()

Call Excel_Exportar_Lsw(lswHistorico)

End Sub

Private Sub btnInformes_Click(Index As Integer)
Dim lngContador As Long, strRuta As String

On Error GoTo vError

If Index = 0 And Len(Trim(txtCedula.Text)) = 0 Then
    Exit Sub
End If


If Not fxValida Then
  Exit Sub
End If

If cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) <> "S" Then
'   MsgBox "La Persona es solo para estado de Asociado (a)", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Personas"
 
 .Connect = glogon.ConectRPT
 
 Select Case Index
      
    Case 0 'Boleta de Registro
        
      .ReportFileName = SIFGlobal.fxPathReportes("Personas_Boleta_Afiliacion.rpt")
      
      .StoredProcParam(0) = Trim(txtCedula)
      .StoredProcParam(1) = 0
    
      .SubreportToChange = "sbBeneficiarios"
      .StoredProcParam(0) = Trim(txtCedula)
        
    Case 1 'Listados
         frmAf_ListadoIngreso.Show vbModal
         Me.MousePointer = vbDefault
         Exit Sub
      
'    Case "socprov"
'       strSQL = "Select Count(*) as Registros From Socios Where EstadoActual='S'"
'       Call OpenRecordSet(rs, strSQL)
'         lngContador = rs!registros
'       rs.Close
'
'      .Formulas(0) = "Socios=" & lngContador
'      .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_SociosProvincia.rpt")
'
'    Case "exsocprov"
'       strSQL = "Select Count(*) as Registros From Socios Where EstadoActual in('A','P')"
'       Call OpenRecordSet(rs, strSQL)
'         lngContador = rs!registros
'       rs.Close
'
'      .Formulas(0) = "Socios=" & lngContador
'      .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_ExSociosProvincia.rpt")
'
'    Case "socup"
'      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_SociosPorUnidad.rpt")
'      .SelectionFormula = "{SOCIOS.ESTADOACTUAL} = 'S'"
'
'    Case "desocup"
'      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_DetalleSociosPorUnidad.rpt")
    
  End Select
   
 .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:

    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnIngresa_Click(Index As Integer)

Dim strTipoMovimiento As String

If Trim(txtCedula) = "" Then Exit Sub

On Error GoTo vError

strTipoMovimiento = ""

Select Case Index
  Case 0, 1
  
'      If Index = 0 Then
'          strTipoMovimiento = "R" 'Re Ingreso
'      Else
'          strTipoMovimiento = "A" 'Activacion
'      End If

'        If Not fxValida Then
'          Exit Sub
'        End If
        
        'Actualiza la Información antes de continuar, y Valida que se actualice
        Call sbSaveRecord
        If Not mSavePass Then
            Exit Sub
        End If

        strSQL = "select COD_MOVIMIENTO from AFI_ESTADOS_CAMBIO" _
               & " where COD_ESTADO in(select EstadoActual from socios where cedula = '" _
               & txtCedula.Text & "') and COD_MOVIMIENTO IN('REI','ACT')"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
         If rs!COD_MOVIMIENTO = "REI" Then strTipoMovimiento = "R"
         If rs!COD_MOVIMIENTO = "ACT" Then strTipoMovimiento = "A"
         rs.MoveNext
        Loop
        rs.Close
  
  
        GLOBALES.gTag = txtInstitucionCod.Text
        GLOBALES.gTag2 = txtCedula & strTipoMovimiento
        GLOBALES.gTag3 = txtApellido1 & " " & txtApellido2 & " " & txtNombre
        frmAF_Reingresos.Show vbModal
        
        GLOBALES.gTag2 = ""
        GLOBALES.gTag3 = ""
        GLOBALES.gTag = ""
  
        Call sbCurrentRecord
  
  
        'Boleta de Afiliacion
        Call btnInformes_Click(0)
        'Contrato Cuenta Sinpe
        Call sbFnd_Contratos_Cuenta_Sinpe(txtCedula.Text)
        
  Case 2 'Ajustar
        GLOBALES.gCedulaActual = txtCedula.Text
        GLOBALES.gTag = cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex)
        frmAF_Ajustes.Show vbModal
        
        Call sbCurrentRecord
End Select



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnNombramiento_Click()
Dim strSQL As String

If Not vEditar Then Exit Sub
If txtCedula.Text = "" Then Exit Sub

On Error GoTo vError


'Revisa Nombramiento / Variaciones para Registrarlas en el Histórico
'Nuevo Modelo: 2020/02/26 {PBN}
strSQL = "exec spAFI_Persona_Nombramientos_Add '" & Trim(txtCedula.Text) & "','" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) _
        & "','" & Format(dtpNombramiento.Value, "yyyy/mm/dd") & "','" & glogon.Usuario & "', 'A'"
Call ConectionExecute(strSQL)

MsgBox "Nombramiento registrado satisfactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnRelacion_Click(Index As Integer)
On Error GoTo vError

Select Case Index
    Case 0 'Registra
        strSQL = "exec spAFI_Persona_Relacion_Add '" & Trim(txtCedula.Text) & "', " & cboR_TipoId.ItemData(cboR_TipoId.ListIndex) _
               & ", '" & Trim(txtR_Cedula.Text) & "', '', '" & Trim(txtR_Apellido1.Text) & "', '" & Trim(txtR_Apellido2.Text) _
               & "', '" & Trim(txtR_Nombre.Text) & "', '" & cboR_TipoVinculo.ItemData(cboR_TipoVinculo.ListIndex) _
               & "', " & IIf((rbCRelacionParentesco(0).Value), 1, 0) & ", '', '', '', '" & glogon.Usuario & "', 'A', " & txtR_Id.Text & ", 1"
        Call OpenRecordSet(rs, strSQL)

        If rs!Pass = 0 Then
            MsgBox rs!ErroMsj, vbExclamation
            Exit Sub
        End If
        
        MsgBox "Registro realizado satisfactoriamente!", vbInformation
        
        Call sbCpl_Relaciones

    Case 1 'Elimina
        strSQL = "exec spAFI_Persona_Relacion_Del " & txtR_Id.Text & ", '" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        MsgBox "Registro Eliminado satisfactoriamente!", vbInformation
        
        Call sbCpl_Relaciones
        
    Case 2 'Nuevo
        txtR_Id.Text = "0"
        txtR_Cedula.Text = ""
        txtR_Apellido1.Text = ""
        txtR_Apellido2.Text = ""
        txtR_Nombre.Text = ""
        
        rbCRelacionParentesco(1).Value = True
        
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnSalarios_Click()
Dim strSQL As String

If Not vEditar Then Exit Sub
If txtCedula.Text = "" Then Exit Sub

On Error GoTo vError

'Registra el Salario + Perfil Transaccional
If CCur(txtSalarioDevengado.Text) > 0 Then

    strSQL = "exec spAFI_Persona_Salarios_Add '" & Trim(txtCedula.Text) & "','" & cboSalarioTipo.ItemData(cboSalarioTipo.ListIndex) _
            & "','" & cboSalarioDivisa.ItemData(cboSalarioDivisa.ListIndex) & "','" & Format(dtpSalarioFecha.Value, "yyyy/mm/dd") _
            & "'," & CCur(txtSalarioDevengado.Text) & "," & CCur(txtSalarioRebajos.Text) & "," & CCur(txtSalarioNeto.Text) _
            & ",'" & IIf(chkSalarioEmbargos.Value = xtpChecked, "S", "N") & "','" & glogon.Usuario & "', 'A'"
    
    strSQL = strSQL & Space(10) & " exec spAFI_Persona_Ingresos_Economicos_Add '" & Trim(txtCedula.Text) _
            & "', " & CCur(txtSalarioDevengado.Text) & ", '" & glogon.Usuario & "', 1"
    Call ConectionExecute(strSQL)
    
    
    MsgBox "Salario registrado satisfactoriamente!", vbInformation

End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnTraDireccion_Click()
Dim vMensaje As String

On Error GoTo vError

If vEditar Then

    vMensaje = ""
    If Trim(cboTraProvincia.Text) = "" Then vMensaje = vMensaje & " - No se especificó la Provincia de la Dirección de Trabajo" & vbCrLf
    If Trim(cboTraCanton.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Cantón de la Dirección de Trabajo" & vbCrLf
    If Trim(cboTraDistrito.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Distrito en la Dirección de Trabajo" & vbCrLf
    
    If Not fxDireccion_Valida(Trim(txtTraDireccion), "-,#,*") Then vMensaje = vMensaje & " - La Dirección Exacta de Trabajo no fue suministrada correctamente!" & vbCrLf

   If Len(vMensaje) = 0 Then
   
     strSQL = "exec spAFI_Persona_Direccion_Add '" & txtCedula.Text & "', '" & cboTraProvincia.ItemData(cboTraProvincia.ListIndex) _
            & "', '" & cboTraCanton.ItemData(cboTraCanton.ListIndex) & "', '" & cboTraDistrito.ItemData(cboTraDistrito.ListIndex) _
            & "', '" & txtTraDireccion.Text & "', '', '', '', '',  '" & glogon.Usuario & "', 'A', 'ProGrX' , 2"
     Call ConectionExecute(strSQL)
 
     MsgBox "Dirección de Trabajo registrada satisfactoriamente!", vbInformation
    
   Else
     MsgBox vMensaje, vbExclamation
   End If
    
End If

Exit Sub

vError:
  
End Sub

Private Sub cboCanton_Click()

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "

End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboSexo.SetFocus
End Sub


Private Sub cboEstadoPersona_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtBoleta.SetFocus
End Sub


Private Sub cboNivelAcademico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPuestoDesc.SetFocus
End Sub

Private Sub cboTraProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboTraProvincia.ItemData(cboTraProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboTraCanton, strSQL, False, True)
vPaso = False

Call cboTraCanton_Click

End Sub

Private Sub cboTraProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTraCanton.SetFocus
End Sub


Private Sub cboTraCanton_Click()
If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboTraProvincia.ItemData(cboTraProvincia.ListIndex) _
            & "' and Canton = '" & cboTraCanton.ItemData(cboTraCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboTraDistrito, strSQL, False, True)

End Sub


Private Sub cboTraCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTraDistrito.SetFocus
End Sub

Private Sub cboTraDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTraDireccion.SetFocus
End Sub


Private Sub chkAportePatronalAdministra_Click()
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Persona_Indicadores '" & Trim(txtCedula.Text) & "', '45', " & chkAportePatronalAdministra.Value & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If chkAportePatronalAdministra.Value = vbUnchecked Then
  Call Bitacora("Aplica", "Activa Administracion de Aporte Patronal a Ced." & txtCedula)
Else
  Call Bitacora("Aplica", "Desactiva Administracion de Aporte Patronal a Ced" & txtCedula)
End If
        
Me.MousePointer = vbDefault
MsgBox "Cambio realizado satisfactoriamente!", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkBienes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeCedula.SetFocus
End Sub

Private Sub chkBloqueo_Click()
On Error GoTo vError

If vPaso Then Exit Sub

Call sbCrdBloqueoCreditos(txtCedula.Text, chkBloqueo.Value)
       
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkConsentimiento_Click()

Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Persona_Indicadores '" & Trim(txtCedula.Text) & "', '29', " & chkConsentimiento.Value & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If chkConsentimiento.Value = vbChecked Then
  Call Bitacora("Aplica", "Firma Consentimiento Informado a Ced." & txtCedula)

        With frmContenedor.Crt
             .Reset
             .WindowShowRefreshBtn = True
             .WindowShowPrintSetupBtn = True
             .WindowState = crptMaximized
             .WindowShowSearchBtn = True
             .WindowTitle = "Reportes Módulo de Personas"
             
             .Connect = glogon.ConectRPT
             
             
             .ReportFileName = SIFGlobal.fxPathReportes("Personas_ConsentimientoInfo.rpt")
             .SelectionFormula = "{SOCIOS.CEDULA} = '" & txtCedula & "'"
        
            
              .PrintReport
        End With

Else
  Call Bitacora("Aplica", "Desactiva Firma Consentimiento Informado a Ced." & txtCedula)
End If
        
Me.MousePointer = vbDefault
MsgBox "Cambio realizado satisfactoriamente!", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub chkDesactivaAporte_Click()
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Persona_Indicadores '" & Trim(txtCedula.Text) & "', '20', " & chkDesactivaAporte.Value & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If chkDesactivaAporte.Value = vbUnchecked Then
  Call Bitacora("Aplica", "Activa Calculo Aportes a Ced." & txtCedula)
Else
  Call Bitacora("Aplica", "Desactiva Calculo Aportes a Ced" & txtCedula)
End If
        
Me.MousePointer = vbDefault
MsgBox "Cambio realizado satisfactoriamente!", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub chkDobleDeduccion_Click()
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Persona_Indicadores '" & Trim(txtCedula.Text) & "', '21', " & chkDobleDeduccion.Value & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
        
If chkDobleDeduccion.Value = vbChecked Then
  Call Bitacora("Aplica", "Activa de Doble Deduc. a Ced." & txtCedula)
Else
  Call Bitacora("Aplica", "Desactivacion de Doble Deduc. a Ced." & txtCedula)
End If
        
Me.MousePointer = vbDefault
MsgBox "Cambio realizado satisfactoriamente!", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkPersonaPolitica_Click()
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Len(txtC_CargoPolitico.Text) <= 5 Then
    MsgBox "Indique el Cargo Político", vbExclamation
    Exit Sub
End If
If dtpC_CargoInicio.Value = dtpC_CargoCorte.Value Then
    MsgBox "Indique el Rango del Nombramiento del Cargo Politico", vbExclamation
    Exit Sub
End If


Me.MousePointer = vbHourglass


strSQL = "exec spAFI_Persona_Indicadores '" & Trim(txtCedula.Text) & "', '46', " & chkPersonaPolitica.Value & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
        
If chkPersonaPolitica.Value = vbChecked Then
  Call Bitacora("Aplica", "Activa Persona Politicamente Expuesta (SUGEF) a Ced." & txtCedula)
Else
  Call Bitacora("Aplica", "Desactivacion Persona Politicamente Expuesta (SUGEF) a Ced." & txtCedula)
End If
        
Me.MousePointer = vbDefault
MsgBox "Cambio realizado satisfactoriamente!", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarDeduciones_Change()
Dim rs As New ADODB.Recordset
Dim vCodigo As String, vColumna As String


On Error GoTo vError

If Not IsNumeric(txtDeduccionesCod.Text) Then
    vCodigo = 0
Else
    vCodigo = txtDeduccionesCod.Text
End If

vColumna = "COD_DEDUCTORA"


If vScroll Then
    strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
           & " from vAFI_Deductoras Where COD_INSTITUCION = " & txtInstitucionCod.Text
    
    If FlatScrollBarDeduciones.Value = 1 Then
       strSQL = strSQL & " and " & vColumna & " > " & vCodigo & " order by " & vColumna & " asc"
    Else
       strSQL = strSQL & " and " & vColumna & " < " & vCodigo & " order by " & vColumna & " desc"
    End If
    
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtDeduccionesCod.Text = rs!Codigo
      txtDeduccionesDesc.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
FlatScrollBarDeduciones.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarPromotor_Change()
Dim rs As New ADODB.Recordset
Dim vCodigo As String, vColumna As String


On Error GoTo vError

If Not IsNumeric(txtPromotorCod.Text) Then
    vCodigo = 0
Else
    vCodigo = txtPromotorCod.Text
End If

vColumna = "ID_PROMOTOR"


If vScroll Then
    strSQL = "select Top 1 " & vColumna & " as 'Codigo',Nombre as 'Descripcion'" _
           & " from Promotores"
    
    If FlatScrollBarPromotor.Value = 1 Then
       strSQL = strSQL & " where " & vColumna & " > " & vCodigo & " order by " & vColumna & " asc"
    Else
       strSQL = strSQL & " where " & vColumna & " < " & vCodigo & " order by " & vColumna & " desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtPromotorCod.Text = rs!Codigo
      txtPromotorDesc.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
FlatScrollBarPromotor.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarRL_Change(Index As Integer)
Dim rs As New ADODB.Recordset
Dim vCodigo As String, vColumna As String, vChar As String, vFiltroAdd As String
Dim txtCodigo As Object, txtDesc As Object

On Error GoTo vError

If Not vScroll Then Exit Sub

vChar = "'"
vFiltroAdd = ""

Select Case Index
   Case 0 'Instituciones
        If Not IsNumeric(txtInstitucionCod.Text) Then
            vCodigo = 0
        Else
            vCodigo = txtInstitucionCod.Text
        End If
        
        vColumna = "COD_INSTITUCION"
        vChar = ""
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from Instituciones"
        
        Set txtCodigo = txtInstitucionCod
        Set txtDesc = txtInstitucionDesc
    
    Case 1 'Departamentos
        vCodigo = txtDeptCodigo.Text
        
        If GLOBALES.SysASEVersion Then
                vColumna = "CODIGO"
               
                strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
                       & " from UPROGRAMATICA"
        
        Else
        
                vColumna = "COD_DEPARTAMENTO"
                vFiltroAdd = " AND COD_INSTITUCION = " & txtInstitucionCod.Text
                
                strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
                       & " from AFDepartamentos"
        End If
        
        Set txtCodigo = txtDeptCodigo
        Set txtDesc = txtDeptDesc
        
    Case 2 'Secciones
        vCodigo = txtSecCodigo.Text
        
        If GLOBALES.SysASEVersion Then
            vColumna = "UT_CODIGO"
            
            strSQL = "select Top 1 " & vColumna & " as 'Codigo',UT_DESCRIPCION as 'Descripcion'" _
                   & " from UTRABAJO"
        Else
            vColumna = "COD_SECCION"
            vFiltroAdd = " AND COD_INSTITUCION = " & txtInstitucionCod.Text & " AND COD_DEPARTAMENTO = '" & txtDeptCodigo.Text & "'"
            
            strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
                   & " from AFSECCIONES"
        End If
        
        Set txtCodigo = txtSecCodigo
        Set txtDesc = txtSecDesc

    Case 3 'Profesion
        If Not IsNumeric(txtProfesionCod.Text) Then
            vCodigo = 0
        Else
            vCodigo = txtProfesionCod.Text
        End If
        
        vColumna = "COD_PROFESION"
        vChar = ""
        vFiltroAdd = " AND ACTIVO = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from AFI_PROFESIONES"
        
        Set txtCodigo = txtProfesionCod
        Set txtDesc = txtProfesionDesc
    
    Case 4 'Sector
        If Not IsNumeric(txtSectorCod.Text) Then
            vCodigo = 0
        Else
            vCodigo = txtSectorCod.Text
        End If
        
        vColumna = "COD_SECTOR"
        vChar = ""
        vFiltroAdd = " AND ACTIVO = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from AFI_SECTORES"
        
        Set txtCodigo = txtSectorCod
        Set txtDesc = txtSectorDesc
    
    
    Case 5 'Tipos de Sociedades
        vCodigo = txtTipoSociedadCod.Text
        vColumna = "COD_SOCIEDAD"
        vFiltroAdd = " AND ACTIVA = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from AFI_SOCIEDADES_TIPOS"
        
        Set txtCodigo = txtTipoSociedadCod
        Set txtDesc = txtTipoSociedadDesc
    
    Case 6 'Actividades Económicas
        vCodigo = txtActividadCod.Text
        vColumna = "COD_ACTIVIDAD"
        vFiltroAdd = " AND ACTIVA = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from AFI_ACTIVIDADES_ECO"
        
        Set txtCodigo = txtActividadCod
        Set txtDesc = txtActividadDesc


    Case 7 'Centros de Trabajo
        vCodigo = txtDeptCodigo.Text
        
        If GLOBALES.SysASEVersion Then
                vColumna = "CODIGO"
               
                strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
                       & " from UPROGRAMATICA"
        
        Else
                vColumna = "COD_DEPARTAMENTO"
                vFiltroAdd = " AND COD_INSTITUCION = " & txtInstitucionCod.Text
                
                strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
                       & " from AFDepartamentos"
        End If
        
        Set txtCodigo = txtCT
        Set txtDesc = txtCT_Desc

End Select

If vScroll Then
    
    If FlatScrollBarRL(Index).Value = 1 Then
       strSQL = strSQL & " where " & vColumna & " > " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " asc"
    Else
       strSQL = strSQL & " where " & vColumna & " < " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Codigo
      txtDesc.Text = rs!Descripcion
      
      If Index = 0 Then
            txtDeduccionesCod.Text = txtInstitucionCod.Text
            txtDeduccionesDesc.Text = txtInstitucionDesc.Text
      End If
    End If
    rs.Close
End If



vScroll = False
FlatScrollBarRL(Index).Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lblDimex_Click()
 
If vEditar Then
 gbDimeX.Visible = True
End If
End Sub

Private Sub lswCumplimiento_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError


strSQL = "exec spAFI_Persona_Productos_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
       & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswDireccion_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

txtDir_Id.Text = Item.Tag

txtDir_Email1.Text = Item.SubItems(5)
txtDir_Email2.Text = Item.SubItems(6)

txtDir_Provincia.Text = Item.SubItems(1)
txtDir_Canton.Text = Item.SubItems(2)
txtDir_Distrito.Text = Item.SubItems(3)
txtDir_Direccion.Text = Item.SubItems(4)

txtDir_Telefono1.Text = Item.SubItems(7)
txtDir_Telefono2.Text = Item.SubItems(8)

txtDir_Usuario.Text = Item.SubItems(9)
txtDir_Fecha.Text = Item.Text

Exit Sub

vError:

End Sub

Private Sub lswHistorico_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

Select Case TituloOpciones.Tag
  
  Case "Motivo"
    strSQL = "exec spAFI_Persona_Motivos_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case "Canal"
    strSQL = "exec spAFI_Persona_Canales_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  Case "Preferencias"
    strSQL = "exec spAFI_Persona_Preferencias_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case "Bienes"
    strSQL = "exec spAFI_Persona_Bienes_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case "Escolaridad"
    strSQL = "exec spAFI_Persona_Escolaridad_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  
  
  Case Else
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub lswRelacion_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If lswRelacion.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

strSQL = "exec spAFI_Persona_Relacion_List '" & txtCedula.Text & "', 1, " & Item.Text
Call OpenRecordSet(rs, strSQL)

txtR_Id.Text = rs!Pr_Id
txtR_Cedula.Text = rs!Cedula & ""
txtR_Apellido1.Text = rs!Apellido1 & ""
txtR_Apellido2.Text = rs!Apellido2 & ""
txtR_Nombre.Text = rs!Nombre & ""

If rs!Empleado = 1 Then
    rbCRelacionParentesco(0).Value = True
Else
    rbCRelacionParentesco(1).Value = True
End If

Call sbCboAsignaDato(cboR_TipoVinculo, rs!Tipo_Relacion_Desc, False, rs!Tipo_Relacion)
Call sbCboAsignaDato(cboR_TipoId, rs!Tipo_Id_Desc, False, rs!Tipo_Id)

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub scrBar_Persona_Change()
Dim rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cedula from socios"
    
    If scrBar_Persona.Value = 1 Then
       strSQL = strSQL & " where cedula > '" & txtCedula.Text & "' and tipo_id = " & cboTipoId.ItemData(cboTipoId.ListIndex) & " order by cedula asc"
    Else
       strSQL = strSQL & " where cedula < '" & txtCedula.Text & "' and tipo_id = " & cboTipoId.ItemData(cboTipoId.ListIndex) & " order by cedula desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCedula.Text = rs!Cedula
      Call txtCedula_LostFocus
    End If
    rs.Close
End If

vScroll = False
scrBar_Persona.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False


Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lblOficina.Caption = ""

vPaso = True

strSQL = "select Catalogo_Id, Descripcion " _
       & " from AFI_CATALOGOS Where Tipo_Id = 8 order by Descripcion"
Call OpenRecordSet(rs, strSQL)

lswCumplimiento.ListItems.Clear
Do While Not rs.EOF
    Set itmX = lswCumplimiento.ListItems.Add(, , rs!Descripcion)
        itmX.Tag = rs!Catalogo_Id
 rs.MoveNext
Loop
rs.Close

vPaso = False


strSQL = "select cod_divisa as 'IdX', Descripcion as 'ItmX' from vSys_Divisas" _
       & " order by divisa_local desc, Descripcion asc"
Call sbCbo_Llena_New(cboSalarioDivisa, strSQL, False, True)

strSQL = "select cod_Pais as 'IdX', Descripcion as 'ItmX' from Paises" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboPaisNac, strSQL, False, True)

strSQL = " select Catalogo_Id as 'IdX', Descripcion as 'ItmX' " _
       & " from AFI_CATALOGOS Where Tipo_Id = 6  and  isnumeric(Catalogo_Id) = 1 order by Descripcion"
Call sbCbo_Llena_New(cboActividad, strSQL, False, True)

strSQL = " select Catalogo_Id as 'IdX', Descripcion as 'ItmX' " _
       & " from AFI_CATALOGOS Where Tipo_Id = 3 order by Descripcion"
Call sbCbo_Llena_New(cboNivelAcademico, strSQL, False, True)


strSQL = "select cod_nacionalidad as 'IdX', Descripcion as 'ItmX' from Sys_nacionalidades" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboNacionalidad, strSQL, False, True)

strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstado, strSQL, False, True)

strSQL = "select Estado_Laboral as 'IdX', Descripcion as 'ItmX' from AFI_ESTADO_LABORAL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoLaboral, strSQL, False, True)


'Opciones Limpias
cboEstadoLaboral.AddItem ""
cboNivelAcademico.AddItem ""
cboActividad.AddItem ""


cboEstadoPersona.Clear

'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)

    'Cumplimiento
    Call sbCbo_Copia(cboTipoId, cboR_TipoId)
vPaso = False

strSQL = " select Catalogo_Id as 'IdX', Descripcion as 'ItmX' " _
       & " from AFI_CATALOGOS Where Tipo_Id = 7 order by Descripcion"
Call sbCbo_Llena_New(cboR_TipoVinculo, strSQL, False, True)


vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
    
    Call sbCbo_Copia(cboProvincia, cboTraProvincia)
vPaso = False


'Carga Parametros
strSQL = "select cod_parametro,valor from afi_parametros where cod_parametro between '01' and '07'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Select Case Trim(rs!Cod_Parametro)
   Case "01"
        vParametros.LargoCedula = CInt(rs!Valor)
   Case "02"
        vParametros.SolicitarTelefonos = IIf((Trim(rs!Valor) = "S"), True, False)
   Case "03"
        vParametros.SolicitarCuentas = IIf((Trim(rs!Valor) = "S"), True, False)
   Case "04"
        vParametros.SoliciarBeneficiario = IIf((Trim(rs!Valor) = "S"), True, False)
   Case "05"
        vParametros.VerificaNombre = IIf((Trim(rs!Valor) = "S"), True, False)
   Case "06"
        vParametros.VerificaPadron = IIf((Trim(rs!Valor) = "S"), True, False)
   Case "07"
        vParametros.BitacoraEspecial = IIf((Trim(rs!Valor) = "S"), True, False)
 End Select
 rs.MoveNext
Loop
rs.Close


Call cboTipoId_Click

Exit Sub

vError:
 


End Sub

Private Sub tpMain_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Call sbTaskPanel_Accion(Item.Id)
End Sub

Private Sub txtActividadCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtActividadDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_actividad,descripcion from AFI_ACTIVIDADES_ECO"
  gBusquedas.Filtro = " and activa = 1"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtActividadCod.Text = Trim(gBusquedas.Resultado)
    txtActividadDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtActividadDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_actividad,descripcion from AFI_ACTIVIDADES_ECO"
  gBusquedas.Filtro = " and activa = 1"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtActividadCod.Text = Trim(gBusquedas.Resultado)
    txtActividadDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub



Private Sub txtAlbaceaCedula_LostFocus()
On Error GoTo vError
        
'Consulta Padron
Call gBase_Padron(txtAlbaceaCedula.Text, "General", rs, "CRI")

If rs.RecordCount > 0 Then
   txtAlbaceaNombre.Text = Trim(rs!Apellido_1) & " " & Trim(rs!Apellido_2) & " " & Trim(rs!Nombre)
End If

vError:

End Sub

Private Sub txtConyugeCedula_LostFocus()
On Error GoTo vError
        
'Consulta Padron
Call gBase_Padron(txtConyugeCedula.Text, "General", rs, "CRI")

If rs.RecordCount > 0 Then
   txtConyugeNombre.Text = Trim(rs!Apellido_1) & " " & Trim(rs!Apellido_2) & " " & Trim(rs!Nombre)
End If

vError:
End Sub

Private Sub txtConyugeTelCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelTrabajoExt.SetFocus
End Sub



Private Sub txtCT_Desc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboNivelAcademico.SetFocus
End Sub

Private Sub txtCT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCT_Desc.SetFocus

If KeyCode = vbKeyF4 Then

 If GLOBALES.SysASEVersion Then
    gBusquedas.Columna = "Codigo"
    gBusquedas.Orden = "Codigo"
    gBusquedas.Consulta = "select Codigo,descripcion from UProgramatica"
 
 Else
    gBusquedas.Columna = "cod_departamento"
    gBusquedas.Orden = "cod_departamento"
    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
    gBusquedas.Filtro = " and cod_institucion = " & txtInstitucionCod.Text
  End If
  
  frmBusquedas.Show vbModal
  txtCT.Text = gBusquedas.Resultado
  txtCT_Desc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtDeduccionesCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeduccionesDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "DESCRIPCION"
   gBusquedas.Orden = "DESCRIPCION"
   gBusquedas.Col1Name = "Código"
   gBusquedas.Col2Name = "Descripción"
 
   gBusquedas.Consulta = "select COD_DEDUCTORA,DESCRIPCION, DESC_CORTA from vAFI_Deductoras"
   gBusquedas.Filtro = " and ACTIVA = 1 and COD_INSTITUCION = " & txtInstitucionCod.Text
   frmBusquedas.Show vbModal
   txtDeduccionesCod.Text = Trim(gBusquedas.Resultado)
   txtDeduccionesDesc.Text = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtDeduccionesDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProfesionCod.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Código"
  gBusquedas.Col2Name = "Descripción"


   gBusquedas.Columna = "DESCRIPCION"
   gBusquedas.Orden = "DESCRIPCION"
   gBusquedas.Consulta = "select COD_DEDUCTORA,DESCRIPCION, DESC_CORTA from vAFI_Deductoras"
   gBusquedas.Filtro = " and ACTIVA = 1 AND COD_INSTITUCION = " & txtInstitucionCod.Text
   frmBusquedas.Show vbModal
   txtDeduccionesCod.Text = Trim(gBusquedas.Resultado)
   txtDeduccionesDesc.Text = Trim(gBusquedas.Resultado2)

End If
End Sub


Private Sub txtInstitucionCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtInstitucionDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "cod_institucion"
  gBusquedas.Orden = "cod_institucion"
  gBusquedas.Consulta = "select cod_institucion,descripcion,desc_Corta from Instituciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtInstitucionCod.Text = Trim(gBusquedas.Resultado)
    txtInstitucionDesc.Text = gBusquedas.Resultado2
  
    txtDeduccionesCod.Text = txtInstitucionCod.Text
    txtDeduccionesDesc.Text = txtInstitucionDesc.Text
  
  End If
End If
End Sub


Private Sub txtInstitucionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Código"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Col3Name = "Desc.Corta"

  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_institucion,descripcion,desc_Corta from Instituciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtInstitucionCod.Text = Trim(gBusquedas.Resultado)
    txtInstitucionDesc.Text = gBusquedas.Resultado2
  
    txtDeduccionesCod.Text = txtInstitucionCod.Text
    txtDeduccionesDesc.Text = txtInstitucionDesc.Text
  
  End If
End If
End Sub


Private Sub txtNumeroPagos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
     Call sbTaskPanel_Accion(Id_TaskItem_Redes)
End If
End Sub

Private Sub txtProfesionCod_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProfesionDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_profesion,descripcion from AFI_Profesiones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtProfesionCod.Text = Trim(gBusquedas.Resultado)
    txtProfesionDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub


Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click

End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub



Private Sub txtProfesionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSectorCod.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_profesion,descripcion from AFI_Profesiones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtProfesionCod.Text = Trim(gBusquedas.Resultado)
    txtProfesionDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub




Private Sub txtPuestoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProfesionCod.SetFocus
End Sub

Private Sub txtR_Cedula_LostFocus()
        
On Error GoTo vError
        
'Consulta Padron
Call gBase_Padron(txtR_Cedula.Text, "General", rs, "CRI")

If rs.RecordCount > 0 Then
   txtR_Apellido1.Text = Trim(rs!Apellido_1)
   txtR_Apellido2.Text = Trim(rs!Apellido_2)
   txtR_Nombre.Text = Trim(rs!Nombre)
End If
        

vError:
End Sub

Private Sub txtSalarioDevengado_GotFocus()
On Error GoTo vError

txtSalarioDevengado.Text = CCur(txtSalarioDevengado.Text)

vError:
End Sub

Private Sub txtSalarioDevengado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSalarioRebajos.SetFocus
End Sub

Private Sub txtSalarioDevengado_LostFocus()
On Error GoTo vError

txtSalarioDevengado.Text = Format(CCur(txtSalarioDevengado.Text), "Standard")

txtSalarioNeto.Text = Format(CCur(txtSalarioDevengado.Text) - CCur(txtSalarioRebajos.Text), "Standard")

vError:
End Sub


Private Sub txtSalarioNeto_GotFocus()
On Error GoTo vError

txtSalarioNeto.Text = CCur(txtSalarioNeto.Text)

vError:
End Sub

Private Sub txtSalarioNeto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkSalarioEmbargos.SetFocus

End Sub

Private Sub txtSalarioNeto_LostFocus()
On Error GoTo vError

txtSalarioNeto.Text = Format(CCur(txtSalarioNeto.Text), "Standard")
txtSalarioDevengado.Text = Format(CCur(txtSalarioNeto.Text) + CCur(txtSalarioRebajos.Text), "Standard")

vError:
End Sub

Private Sub txtSalarioRebajos_GotFocus()
On Error GoTo vError

txtSalarioRebajos.Text = CCur(txtSalarioRebajos.Text)

vError:
End Sub

Private Sub txtSalarioRebajos_LostFocus()
On Error GoTo vError

txtSalarioRebajos.Text = Format(CCur(txtSalarioRebajos.Text), "Standard")

txtSalarioNeto.Text = Format(CCur(txtSalarioDevengado.Text) - CCur(txtSalarioRebajos.Text), "Standard")

vError:
End Sub

Private Sub txtSectorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSectorDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_sector,descripcion from AFI_Sectores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtSectorCod.Text = Trim(gBusquedas.Resultado)
    txtSectorDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub

Private Sub cboSexo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpNacimiento.SetFocus
End Sub


Private Sub cboTipoId_Click()
If vPaso Then Exit Sub

'Call sbClearControles

If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
    fraTipo.Visible = True
Else
    fraTipo.Visible = False
End If


End Sub



Private Sub cmdCambioPriDeduc_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update socios set Prideduc = " & txtPrideduc.Text _
       & " where cedula = '" & txtCedula & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "1er Deduccion de Aportes a Ced." & txtCedula)
MsgBox "Modificacion de la 1er Deducción de Aportes Patronales realizada Satisfactoriamente...", vbInformation
        
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub cmdNotasAdv_Click()
Dim strSQL As String

On Error GoTo vError

txtNotasAdv.Text = fxSysCleanTxtInject(txtNotasAdv.Text)

If Len(txtNotasAdv.Text) < 20 Then
    MsgBox "La nota para el bloqueo no es valida, favor agregar más información!", vbExclamation
    Exit Sub
End If

strSQL = "exec spAFI_Persona_Indicadores '" & Trim(txtCedula.Text) & "', '22', 0, '" & glogon.Usuario & "', '" & txtNotasAdv.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Notas de Advertencia Ced." & txtCedula)

MsgBox "NOTA REGISTRADA...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
Dim fraX As Frame
Dim rs As New ADODB.Recordset

On Error GoTo vError

vModulo = 1

'Ajuste de Alto
Me.Height = 9735

gbDimeX.Left = TituloOpcion.Left
gbDimeX.top = TituloOpcion.top
gbDimeX.Visible = False

Call Formularios(Me)
 
Call sbTaskPanel_Load
 
 vScroll = False
 scrBar_Persona.Value = 0
 vScroll = True
 
'gbOpciones.Visible = False

vEditar = False

dtpFechaIngreso.Value = fxFechaServidor
dtpNacimiento.Value = dtpFechaIngreso.Value
dtpCedulaVence.Value = dtpFechaIngreso.Value

vFechaActual = dtpFechaIngreso.Value



cboSexo.AddItem "Masculino"
cboSexo.AddItem "Femenino"
cboSexo.AddItem "Otro"
cboSexo.Text = "Masculino"


cboPatrono.AddItem "Público"
cboPatrono.ItemData(cboPatrono.ListCount - 1) = "0"
cboPatrono.AddItem "Privado"
cboPatrono.ItemData(cboPatrono.ListCount - 1) = "1"
cboPatrono.AddItem "Otro"
cboPatrono.ItemData(cboPatrono.ListCount - 1) = "2"
cboPatrono.Text = "Público"

'Revisa cual Tipo de Identificacion es Juridica (Solo es Valido la Primera)
vTipoJuridica = 0
strSQL = "select TIPO_ID from AFI_TIPOS_IDS where Tipo_Personeria = 'J'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    vTipoJuridica = rs!Tipo_Id
End If
rs.Close

cboSalarioTipo.Clear
cboSalarioTipo.AddItem "Colilla"
cboSalarioTipo.ItemData(cboSalarioTipo.ListCount - 1) = "C"
cboSalarioTipo.AddItem "Constancia"
cboSalarioTipo.ItemData(cboSalarioTipo.ListCount - 1) = "X"
cboSalarioTipo.Text = "Colilla"

Call sbBarra_Accion("nuevo")

Call sbLockControles("L")
Call RefrescaTags(Me)


With lswCumplimiento.ColumnHeaders
    .Clear
    .Add , , "", lswCumplimiento.Width - 150
End With

btnIngresa.Item(0).Enabled = False 'Ingreos
'btnIngresa.Item(1).Enabled = False 'Activacion


tcMain.Item(0).Selected = True
tcAux.Item(0).Selected = True

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




Private Sub sbLimpiaDatos()
Dim i As Integer, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)

'Carga Estados de la Persona
strSQL = "select RTRIM(E.COD_ESTADO) as 'idX' , rtrim(E.DESCRIPCION) AS itmX" _
       & " from AFI_ESTADOS_PERSONA E inner join AFI_ESTADOS_CAMBIO C on E.COD_ESTADO = C.COD_ESTADO" _
       & " where C.COD_MOVIMIENTO = 'ING' and E.ACTIVO = 1 Group by E.COD_ESTADO, E.DESCRIPCION"
cboEstadoPersona.Clear
Call sbCbo_Llena_New(cboEstadoPersona, strSQL, False, True)

  
 gbDimeX.Visible = False

If vEditar Then
    txtDeduccionesCod.Enabled = True
    txtDeduccionesDesc.Enabled = True
    FlatScrollBarDeduciones.Enabled = True
Else
    txtDeduccionesCod.Enabled = False
    txtDeduccionesDesc.Enabled = False
    FlatScrollBarDeduciones.Enabled = False
End If

dtpFechaIngreso.Value = vFechaActual
dtpNacimiento.Value = vFechaActual
dtpNombramiento.Value = vFechaActual

dtpSalarioFecha.Value = vFechaActual

txtNombreComercial.Text = ""
txtRazonSocial.Text = ""

txtNombre.Text = ""
txtApellido1.Text = ""
txtApellido2.Text = ""

txtHijos.Text = 0
txtNumeroPagos.Text = 2
txtBoleta.Text = 0

chkBeneficiarios.Value = xtpChecked
chkTrabajoPropio.Value = xtpUnchecked
cboPatrono.Text = "Público"
txtPuestoDesc.Text = ""

rbCES.Item(0).Value = True 'CES 1

If GLOBALES.SysASEVersion Then
    lblDepartamento.Caption = "U.Programatica"
    lblSeccion.Caption = "U.Trabajo"
End If


chkBienes.Value = vbUnchecked
lblOficina.Caption = ""


cboSexo.Text = "Masculino"
'cboEstado.Text = "Soltero"

txtEmail.Text = ""
txtEmail_02.Text = ""


txtDireccion.Text = ""
txtApartado.Text = ""


cboTraCanton.Clear
cboTraDistrito.Clear
txtTraDireccion.Text = ""


txtSN_Blog.Text = ""
txtSN_Facebook.Text = ""
txtSN_Twitter.Text = ""
txtSN_Instagram.Text = ""
txtSN_LinkedIn.Text = ""

txtConyugeCedula.Text = ""
txtConyugeNombre.Text = ""
txtConyugeTelCelular.Text = ""
txtConyugeTelTrabajo.Text = ""
txtConyugeTelTrabajoExt.Text = ""


txtAlbaceaCedula.Text = ""
txtAlbaceaNombre.Text = ""
txtAlbaceaTelCelular.Text = ""
txtAlbaceaTelTrabajo.Text = ""
txtAlbaceaTelTrabajoExt.Text = ""


txtSalarioDevengado.Text = "0.00"
txtSalarioNeto.Text = "0.00"
txtSalarioRebajos.Text = "0.00"

chkSalarioEmbargos.Value = xtpUnchecked


'Cumplimiento
dtpC_CargoCorte.Value = vFechaActual
dtpC_CargoInicio.Value = vFechaActual

vPaso = True
    chkPersonaPolitica.Value = xtpUnchecked
vPaso = False

txtC_CargoPolitico.Text = ""

rbCES(0).Value = True


strSQL = "exec spAFI_RegistroDefault"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Select Case Trim(rs!Tipo)
   Case "INSTITUCION"
      txtDeduccionesCod.Text = rs!Codigo
      txtDeduccionesDesc.Text = rs!Descripcion
      
      txtInstitucionCod.Text = rs!Codigo
      txtInstitucionDesc.Text = rs!Descripcion
        
   Case "PROFESION"
      txtProfesionCod.Text = rs!Codigo
      txtProfesionDesc.Text = rs!Descripcion
      
      txtProfesionCod.Text = ""
      txtProfesionDesc.Text = ""
      
      
   Case "PROMOTORES"
      txtPromotorCod.Text = rs!Codigo
      txtPromotorDesc.Text = rs!Descripcion
      
   Case "SECTOR"
      txtSectorCod.Text = rs!Codigo
      txtSectorDesc.Text = rs!Descripcion
 End Select
 rs.MoveNext
Loop
rs.Close

vPaso = True

cboNivelAcademico.Text = ""
cboEstadoLaboral.Text = ""
cboActividad.Text = ""

chkConsentimiento.Value = xtpChecked
chkAportePatronalAdministra.Value = xtpChecked

tcMain.Item(0).Selected = True
tcAux.Item(0).Selected = True

vPaso = False

Call cboTipoId_Click

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtAlbaceaCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAlbaceaNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtAlbaceaCedula = gBusquedas.Resultado
  txtAlbaceaNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtAlbaceaNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSN_Facebook.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtAlbaceaCedula = gBusquedas.Resultado
  txtAlbaceaNombre = gBusquedas.Resultado2
End If
End Sub

Private Sub txtApartado_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus

End Sub


Private Sub txtApellido1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido2.SetFocus
End Sub

Private Sub txtApellido1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtApellido2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtApellido2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPromotorCod.SetFocus
End Sub



Private Sub txtCedAlternativa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 If fraTipo.Visible Then
     txtNombreComercial.SetFocus
 Else
     txtNombre.SetFocus
 End If
End If



If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtCedula)
   txtCedula = ""
   vCedula = ""
   
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Id Alterno"
   gBusquedas.Col3Name = "Nombre"
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "CedulaR"
   gBusquedas.Orden = "CedulaR"
   gBusquedas.Consulta = "Select Cedula,CedulaR, Nombre From Socios"
   gBusquedas.Filtro = " and Tipo_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
   frmBusquedas.Show vbModal
   
   txtCedula = Trim(gBusquedas.Resultado)
   txtCedula_LostFocus
End If


End Sub

Private Sub txtCedula_GotFocus()
vSeek = 1
'SSTab1.Tab = 0
'ssTabSubX.Tab = 0

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtCedula)
   txtCedula = ""
   vCedula = ""
   
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Id Alterno"
   gBusquedas.Col3Name = "Nombre"
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "Select Cedula,CedulaR, Nombre From Socios"
   gBusquedas.Filtro = " and Tipo_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)

   frmBusquedas.Show vbModal
   
   txtCedula = Trim(gBusquedas.Resultado)
   txtCedula_LostFocus
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtCedula_LostFocus
End If

End Sub


Private Sub txtCedula_LostFocus()

If Trim(txtCedula) = "" Then
  If vEditar Then
     vEditar = False
     Call sbBarra_Accion("nuevo")
     Call RefrescaTags(Me)
     Call sbClearControles
     Call sbLockControles("L")
  End If
Else
  If Not vEditar Or (vEditar And vCedula <> Trim(txtCedula)) Then
     Call sbCurrentRecord
  End If
End If

End Sub


Private Sub txtPromotorCod_GotFocus()
If txtPromotorCod.Text = "" Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Id_Promotor,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorCod = Trim(gBusquedas.Resultado)
   txtPromotorDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtPromotorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPromotorDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "ID_PROMOTOR"
   gBusquedas.Orden = "ID_PROMOTOR"
   gBusquedas.Consulta = "select Id_Promotor ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorCod = Trim(gBusquedas.Resultado)
   txtPromotorDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtConyugeCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtConyugeCedula = gBusquedas.Resultado
  txtConyugeNombre = gBusquedas.Resultado2
End If

End Sub



Private Sub txtConyugeNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelCelular.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtConyugeCedula = gBusquedas.Resultado
  txtConyugeNombre = gBusquedas.Resultado2
End If

End Sub


Private Sub txtConyugeTelTrabajo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelTrabajoExt.SetFocus
End Sub

Private Sub txtConyugeTelTrabajoExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAlbaceaCedula.SetFocus
End Sub



Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus

If KeyCode = vbKeyF4 Then

  gBusquedas.Col1Name = "Código"
  gBusquedas.Col2Name = "Descripción"

 If GLOBALES.SysASEVersion Then
    gBusquedas.Columna = "Codigo"
    gBusquedas.Orden = "Codigo"
    gBusquedas.Consulta = "select Codigo,descripcion from UProgramatica"
 
 Else
    gBusquedas.Columna = "cod_departamento"
    gBusquedas.Orden = "cod_departamento"
    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
    gBusquedas.Filtro = " and cod_institucion = " & txtInstitucionCod.Text
  End If
  
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If


End Sub

Private Sub txtDeptCodigo_LostFocus()
 txtDeptDesc.Text = fxgAFIDepartamento(txtInstitucionCod.Text, txtDeptCodigo.Text)
End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then

 gBusquedas.Columna = "descripcion"
 gBusquedas.Orden = "descripcion"
 
 gBusquedas.Col1Name = "Código"
 gBusquedas.Col2Name = "Descripción"
 
 If GLOBALES.SysASEVersion Then
    gBusquedas.Consulta = "select Codigo,descripcion from UProgramatica"
 Else
    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
    gBusquedas.Filtro = " and cod_institucion = " & txtInstitucionCod.Text
 End If
 
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If


End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then txtNotificaciones.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail_02.SetFocus
End Sub

Private Sub txtEmail_02_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartado.SetFocus
End Sub


Private Sub txtHijos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkBienes.SetFocus
End Sub

Private Sub txtNombre_GotFocus()
vSeek = 2
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtCedula)
   txtCedula = ""
   vCedula = ""
      
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Id Alterno"
   gBusquedas.Col3Name = "Nombre"
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "Select Cedula,CedulaR, Nombre From Socios"
   gBusquedas.Filtro = " and Tipo_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)

   
   frmBusquedas.Show vbModal
   
   txtCedula = Trim(gBusquedas.Resultado)
   txtCedula_LostFocus
End If

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



Private Sub txtNombreComercial_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRazonSocial.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtCedula)
   txtCedula = ""
   vCedula = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "nombre"
   gBusquedas.Orden = "nombre"
   gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
   gBusquedas.Filtro = " and Tipo_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
   frmBusquedas.Show vbModal
   
   txtCedula = Trim(gBusquedas.Resultado)
   txtCedula_LostFocus
End If
End Sub

Private Sub txtPromotorDesc_GotFocus()

If txtPromotorDesc.Text = "" Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Id_Promotor,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorCod = Trim(gBusquedas.Resultado)
   txtPromotorDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtPromotorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtDeduccionesCod.Enabled = True Then txtDeduccionesCod.SetFocus
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtDeduccionesCod.Enabled = False Then txtEmail.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Id_Promotor,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorCod = Trim(gBusquedas.Resultado)
   txtPromotorDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtNotificaciones_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 Call sbTaskPanel_Accion(Id_TaskItem_RelacionLaboral)
End If
End Sub



Private Sub txtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstadoPersona.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtCedula)
   txtCedula = ""
   vCedula = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Razon_Social"
   gBusquedas.Orden = "Razon_Social"
   gBusquedas.Consulta = "Select Cedula,Razon_Social From Socios"
   gBusquedas.Filtro = " and Tipo_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
   frmBusquedas.Show vbModal
   
   txtCedula = Trim(gBusquedas.Resultado)
   txtCedula_LostFocus
End If
End Sub

Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then

  gBusquedas.Col1Name = "Código"
  gBusquedas.Col2Name = "Descripción"
  
  If GLOBALES.SysASEVersion Then
        gBusquedas.Columna = "UT_Codigo"
        gBusquedas.Orden = "UT_Codigo"
        gBusquedas.Consulta = "select UT_Codigo,UT_descripcion from UTrabajo"
  Else
        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from AFSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & txtInstitucionCod.Text _
                  & " and cod_departamento = '" & txtDeptCodigo & "'"
  End If
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtSecCodigo_LostFocus()
 txtSecDesc.Text = fxgAFISeccion(txtInstitucionCod.Text, txtDeptCodigo.Text, txtSecCodigo.Text)
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCT.SetFocus

If KeyCode = vbKeyF4 Then

  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Col1Name = "Código"
  gBusquedas.Col2Name = "Descripción"
  
  If GLOBALES.SysASEVersion Then
        gBusquedas.Consulta = "select UT_Codigo, UT_Descripcion from UTrabajo"
  Else
        gBusquedas.Consulta = "select cod_seccion,descripcion from AFSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & txtInstitucionCod.Text _
                  & " and cod_departamento = '" & txtDeptCodigo & "'"
  End If
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub


Private Sub txtSectorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  tcTrabajo.Item(0).Selected = True
  txtTipoSociedadCod.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_sector,descripcion from AFI_Sectores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtSectorCod.Text = Trim(gBusquedas.Resultado)
    txtSectorDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtSN_Blog_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)
End If
End Sub

Private Sub txtSN_Facebook_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSN_Twitter.SetFocus

End Sub


Private Sub txtSN_Instagram_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSN_Blog.SetFocus

End Sub

Private Sub txtSN_LinkedIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSN_Instagram.SetFocus

End Sub

Private Sub txtSN_Twitter_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSN_LinkedIn.SetFocus

End Sub

Private Sub txtTipoSociedadCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSociedadDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_sociedad,descripcion from AFI_SOCIEDADES_TIPOS"
  gBusquedas.Filtro = " and activa = 1"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtTipoSociedadCod.Text = Trim(gBusquedas.Resultado)
    txtTipoSociedadDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtTipoSociedadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtActividadCod.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_sociedad,descripcion from AFI_SOCIEDADES_TIPOS"
  gBusquedas.Filtro = " and activa = 1"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtTipoSociedadCod.Text = Trim(gBusquedas.Resultado)
    txtTipoSociedadDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub

Private Sub sbInsertAhorro()

On Error GoTo vError

strSQL = "exec spAFI_PERSONA_PATRIMONIO_Vincula '" & Trim(txtCedula.Text) & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub udPriDeduc_DownClick()
On Error Resume Next

txtPrideduc.Text = fxFechaProcesoAnterior(txtPrideduc.Text)

End Sub

Private Sub udPriDeduc_UpClick()
On Error Resume Next
txtPrideduc.Text = fxFechaProcesoSiguiente(txtPrideduc.Text)

End Sub
