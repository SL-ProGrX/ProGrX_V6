VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Cat_Conceptos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conceptos de Nómina"
   ClientHeight    =   7905
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10050
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   252
      Left            =   8160
      TabIndex        =   0
      Top             =   600
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activo?"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   11880
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
      ItemCount       =   5
      Item(0).Caption =   "Concepto"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "Label1(1)"
      Item(0).Control(1)=   "txtDescripcion"
      Item(0).Control(2)=   "Label1(2)"
      Item(0).Control(3)=   "gbSalario"
      Item(0).Control(4)=   "txtPrioridad"
      Item(0).Control(5)=   "chkPatronoConcepto"
      Item(0).Control(6)=   "chkMuestraENColilla"
      Item(0).Control(7)=   "cboTipoConcepto"
      Item(0).Control(8)=   "Label1(7)"
      Item(0).Control(9)=   "Label1(4)"
      Item(0).Control(10)=   "dtpInicio"
      Item(0).Control(11)=   "Label1(5)"
      Item(0).Control(12)=   "dtpFinaliza"
      Item(0).Control(13)=   "GroupBox1"
      Item(0).Control(14)=   "chkAplicacionGeneral"
      Item(0).Control(15)=   "chkSalarioImbargable"
      Item(0).Control(16)=   "chkNoVence"
      Item(0).Control(17)=   "cboGrupo"
      Item(0).Control(18)=   "Label1(3)"
      Item(1).Caption =   "Tablas"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "vGrid"
      Item(1).Control(1)=   "scConcepto"
      Item(2).Caption =   "Otros"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "gbEntidad"
      Item(2).Control(1)=   "GroupBox2"
      Item(3).Caption =   "Detalle"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "ShortcutCaption1"
      Item(3).Control(1)=   "gDetalle"
      Item(3).Control(2)=   "scDetalle"
      Item(3).Control(3)=   "btnBack(0)"
      Item(4).Caption =   "Exclusión"
      Item(4).ControlCount=   7
      Item(4).Control(0)=   "ShortcutCaption2"
      Item(4).Control(1)=   "txtExcID"
      Item(4).Control(2)=   "txtExcDesc"
      Item(4).Control(3)=   "btnTool_Exc(0)"
      Item(4).Control(4)=   "btnTool_Exc(1)"
      Item(4).Control(5)=   "btnTool_Exc(2)"
      Item(4).Control(6)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4932
         Left            =   -69760
         TabIndex        =   64
         Top             =   1680
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   8700
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtExcDesc 
         Height          =   312
         Left            =   -67840
         TabIndex        =   59
         Top             =   1200
         Visible         =   0   'False
         Width           =   6012
         _Version        =   1441793
         _ExtentX        =   10604
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
      Begin XtremeSuiteControls.GroupBox gbSalario 
         Height          =   2055
         Left            =   360
         TabIndex        =   2
         Top             =   3120
         Width           =   9375
         _Version        =   1441793
         _ExtentX        =   16536
         _ExtentY        =   3625
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnDetalle 
            Height          =   315
            Left            =   6600
            TabIndex        =   52
            ToolTipText     =   "Detalle en Rubros del Valor"
            Top             =   240
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Detalle Rubros"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtAplValor 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   4560
            TabIndex        =   3
            Top             =   240
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
         Begin XtremeSuiteControls.ComboBox cboAplTipo 
            Height          =   312
            Left            =   1440
            TabIndex        =   19
            Top             =   240
            Width           =   1932
            _Version        =   1441793
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboAplBase 
            Height          =   312
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   1932
            _Version        =   1441793
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboAplResume 
            Height          =   312
            Left            =   1440
            TabIndex        =   23
            Top             =   1200
            Width           =   1932
            _Version        =   1441793
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.CheckBox chkPermiteCambios 
            Height          =   252
            Left            =   4560
            TabIndex        =   46
            Top             =   1440
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite Cambios (Manual/Carga de Listados)?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkCreditoFiscal 
            Height          =   252
            Left            =   4560
            TabIndex        =   48
            Top             =   1800
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Crédito Fiscal al Cálculo?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtAplFx 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   555
            Left            =   4560
            TabIndex        =   24
            Top             =   720
            Width           =   4815
            _Version        =   1441793
            _ExtentX        =   8493
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
            Alignment       =   2
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnExclusion 
            Height          =   315
            Left            =   7920
            TabIndex        =   53
            ToolTipText     =   "Excluir del Cálculo Base"
            Top             =   240
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Exclusión"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnActualiza_Personas 
            Height          =   315
            Left            =   6120
            TabIndex        =   65
            ToolTipText     =   "Actualiza los Conceptos Asignados a Empleados"
            Top             =   240
            Width           =   375
            _Version        =   1441793
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
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
            Appearance      =   16
            Picture         =   "frmRH_Cat_Conceptos.frx":0000
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Left            =   3720
            TabIndex        =   25
            Top             =   720
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "f(x)"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   372
            Index           =   10
            Left            =   0
            TabIndex        =   22
            Top             =   1200
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Resume en"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   372
            Index           =   9
            Left            =   0
            TabIndex        =   20
            Top             =   720
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Utiliza como Base"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   372
            Index           =   8
            Left            =   0
            TabIndex        =   18
            Top             =   240
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Tipo Aplicación"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblAplBase 
            Height          =   252
            Left            =   3720
            TabIndex        =   4
            Top             =   240
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Valor"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
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
      Begin XtremeSuiteControls.CheckBox chkPatronoConcepto 
         Height          =   252
         Left            =   1800
         TabIndex        =   7
         Top             =   2400
         Width           =   3012
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Concepto de Registro de Patrono?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkMuestraENColilla 
         Height          =   252
         Left            =   1800
         TabIndex        =   8
         Top             =   2040
         Width           =   3012
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Muestra en Colilla?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboTipoConcepto 
         Height          =   330
         Left            =   1800
         TabIndex        =   15
         Top             =   1200
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
      Begin XtremeSuiteControls.CheckBox chkAplicacionGeneral 
         Height          =   252
         Left            =   1800
         TabIndex        =   17
         Top             =   1680
         Width           =   3012
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplicación General?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtPrioridad 
         Height          =   315
         Left            =   7560
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
         _Version        =   1441793
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   7560
         TabIndex        =   27
         Top             =   1680
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1455
         Left            =   240
         TabIndex        =   30
         Top             =   5160
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Contabilidad:"
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
            Height          =   312
            Left            =   1560
            TabIndex        =   31
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Width           =   1932
            _Version        =   1441793
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   315
            Left            =   3480
            TabIndex        =   32
            Top             =   600
            Width           =   6015
            _Version        =   1441793
            _ExtentX        =   10610
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaCierreId 
            Height          =   312
            Left            =   1560
            TabIndex        =   33
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1080
            Width           =   1932
            _Version        =   1441793
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaCierreDesc 
            Height          =   315
            Left            =   3480
            TabIndex        =   34
            Top             =   1080
            Width           =   6015
            _Version        =   1441793
            _ExtentX        =   10610
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
         Begin XtremeSuiteControls.CheckBox chkNoContabiliza 
            Height          =   252
            Left            =   1560
            TabIndex        =   47
            Top             =   240
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "No Contabiliza (Solo Cálculo)?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.Label lblCuentaCierre 
            Height          =   372
            Left            =   240
            TabIndex        =   36
            Top             =   1080
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cuenta Cierre"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   372
            Index           =   6
            Left            =   240
            TabIndex        =   35
            Top             =   600
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cuenta"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5772
         Left            =   -67360
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   6972
         _Version        =   524288
         _ExtentX        =   12298
         _ExtentY        =   10181
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
         MaxCols         =   485
         ScrollBars      =   2
         SpreadDesigner  =   "frmRH_Cat_Conceptos.frx":0719
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.GroupBox gbEntidad 
         Height          =   1092
         Left            =   -69760
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Concepto vinculado con la Entidad:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtEntidadId 
            Height          =   312
            Left            =   1080
            TabIndex        =   40
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   480
            Width           =   1932
            _Version        =   1441793
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
         Begin XtremeSuiteControls.FlatEdit txtEntidadDesc 
            Height          =   312
            Left            =   3000
            TabIndex        =   41
            Top             =   480
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1092
         Left            =   -69760
         TabIndex        =   42
         Top             =   2040
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Concepto vinculado con Otro Concepto:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtConceptoVinId 
            Height          =   312
            Left            =   1080
            TabIndex        =   43
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   480
            Width           =   1932
            _Version        =   1441793
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
         Begin XtremeSuiteControls.FlatEdit txtConceptoVinDesc 
            Height          =   312
            Left            =   3000
            TabIndex        =   44
            Top             =   480
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
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
      End
      Begin XtremeSuiteControls.CheckBox chkSalarioImbargable 
         Height          =   252
         Left            =   1800
         TabIndex        =   45
         Top             =   2760
         Width           =   3012
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Revisa Salario Inembargable?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFinaliza 
         Height          =   315
         Left            =   7560
         TabIndex        =   29
         Top             =   2160
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.CheckBox chkNoVence 
         Height          =   375
         Left            =   7560
         TabIndex        =   49
         Top             =   2520
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Sin Vencimiento"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboGrupo 
         Height          =   330
         Left            =   1800
         TabIndex        =   50
         Top             =   840
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
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
      Begin FPSpreadADO.fpSpread gDetalle 
         Height          =   5772
         Left            =   -67360
         TabIndex        =   54
         Top             =   840
         Visible         =   0   'False
         Width           =   6972
         _Version        =   524288
         _ExtentX        =   12298
         _ExtentY        =   10181
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmRH_Cat_Conceptos.frx":0DD1
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnBack 
         Height          =   312
         Index           =   0
         Left            =   -61960
         TabIndex        =   57
         ToolTipText     =   "Detalle en Rubros del Valor"
         Top             =   380
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Regresar"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtExcID 
         Height          =   312
         Left            =   -69760
         TabIndex        =   58
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
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
      Begin XtremeSuiteControls.PushButton btnTool_Exc 
         Height          =   312
         Index           =   0
         Left            =   -61840
         TabIndex        =   61
         ToolTipText     =   "Detalle en Rubros del Valor"
         Top             =   1200
         Visible         =   0   'False
         Width           =   372
         _Version        =   1441793
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
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
         Picture         =   "frmRH_Cat_Conceptos.frx":13B6
      End
      Begin XtremeSuiteControls.PushButton btnTool_Exc 
         Height          =   312
         Index           =   1
         Left            =   -61360
         TabIndex        =   62
         ToolTipText     =   "Detalle en Rubros del Valor"
         Top             =   1200
         Visible         =   0   'False
         Width           =   372
         _Version        =   1441793
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
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
         Picture         =   "frmRH_Cat_Conceptos.frx":19E8
      End
      Begin XtremeSuiteControls.PushButton btnTool_Exc 
         Height          =   312
         Index           =   2
         Left            =   -61000
         TabIndex        =   63
         ToolTipText     =   "Detalle en Rubros del Valor"
         Top             =   1200
         Visible         =   0   'False
         Width           =   372
         _Version        =   1441793
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
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
         Picture         =   "frmRH_Cat_Conceptos.frx":2119
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   372
         Left            =   -70000
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441793
         _ExtentX        =   16954
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Indicar los Conceptos que se Restaran a la Base de la Aplicación del Concepto  Madre"
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scDetalle 
         Height          =   372
         Left            =   -70000
         TabIndex        =   56
         Top             =   360
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441793
         _ExtentX        =   16954
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Valor a Detallar: xx, Pendiente X"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   5772
         Left            =   -70000
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
         _ExtentY        =   10181
         _StockProps     =   14
         Caption         =   "Detalle de Rubros:"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   372
         Index           =   3
         Left            =   360
         TabIndex        =   51
         Top             =   840
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Grupo"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scConcepto 
         Height          =   5772
         Left            =   -70000
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
         _ExtentY        =   10181
         _StockProps     =   14
         Caption         =   "Tabla de Cálculo:"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Index           =   5
         Left            =   6000
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Fecha de Finalización"
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
         Alignment       =   5
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   4
         Left            =   6000
         TabIndex        =   26
         Top             =   1680
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Fecha de Inicio"
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
         Alignment       =   5
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   7
         Left            =   6000
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Prioridad"
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
         Alignment       =   5
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripción"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tipo de Concepto"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1920
      TabIndex        =   13
      Top             =   600
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   600
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Concepto"
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmRH_Cat_Conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim vEdita  As Boolean
Dim vCodigo As String, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Function fxExiste(vCodigo As String) As Boolean
strSQL = "select isnull(count(*),0) as 'Existe'" _
       & " from RH_CONCEPTOS where COD_CONCEPTO =  '" & vCodigo & "' "
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  fxExiste = False
Else
  fxExiste = True
End If
rs.Close
End Function

Private Function fxConcepto_Id_New() As String

strSQL = "select isnull(count(*),0)+ 1 as 'Sequencia'" _
       & " from RH_CONCEPTOS"
Call OpenRecordSet(rs, strSQL)

fxConcepto_Id_New = Format(rs!Sequencia, "0000")

rs.Close

End Function

Private Function fxPeriodad_Id_New(pTipo As String) As String

strSQL = "select isnull(count(*),0) + 1 as 'Sequencia'" _
       & " from RH_CONCEPTOS where Prioridad like '" & Trim(pTipo) & "-%' "
Call OpenRecordSet(rs, strSQL)

fxPeriodad_Id_New = Trim(pTipo) & "-" & Format(rs!Sequencia, "0000")

rs.Close

End Function

Private Sub sbDetalle_List()

On Error GoTo vError

tcMain(3).Visible = True
tcMain(3).Selected = True

scDetalle.Caption = "Valor a Detallar: " & txtAplValor.Text

strSQL = "select COD_DETALLE, DESCRIPCION, BASE_MINIMA, VALOR" _
       & " From RH_CONCEPTOS_DETALLE" _
       & " Where cod_Concepto = '" & vCodigo & "'"

Call sbCargaGrid(gDetalle, 4, strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnActualiza_Personas_Click()
On Error GoTo vError

Dim strSQL As String

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Concepto_Persona_Actualiza '" & txtCodigo.Text & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Actualiza", "Concepto a Personas vinculadas: " & txtCodigo.Text)

Me.MousePointer = vbDefault

MsgBox "Concepto Actualizado a todos los empleados vinculados!", vbInformation

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnBack_Click(Index As Integer)
tcMain(0).Selected = True

End Sub

Private Sub btnDetalle_Click()

Call sbDetalle_List

End Sub

Private Sub sbExcluye_List()

On Error GoTo vError

tcMain(4).Visible = True
tcMain(4).Selected = True

lsw.ListItems.Clear

txtExcID.Text = ""
txtExcDesc.Text = ""

strSQL = "exec spRH_Concepto_Excluye_Lista '" & vCodigo & "'"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!COD_CONCEPTO_EXC)
      itmX.SubItems(1) = rs!Descripcion
  rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbExcluye_Registro(pMov As String)

On Error GoTo vError

strSQL = "exec spRH_Concepto_Excluye_Registro '" & vCodigo & "','" & txtExcID.Text _
        & "','" & glogon.Usuario & "', '" & pMov & "'"
Call ConectionExecute(strSQL)

Call sbExcluye_List

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnExclusion_Click()

Call sbExcluye_List

End Sub

Private Sub btnTool_Exc_Click(Index As Integer)

Select Case Index
  Case 0 'Nuevo
    txtExcID.Text = ""
    txtEntidadDesc.Text = ""
  
  Case 1 'Guardar
    Call sbExcluye_Registro("A")
  
  Case 2 'Eliminar
    Call sbExcluye_Registro("E")
  
End Select

End Sub

Private Sub cboAplTipo_Click()

If vPaso Then Exit Sub

tcMain(0).Selected = True
tcMain(1).Visible = False

txtAplFx.Locked = True

Select Case Mid(cboAplTipo.Text, 1, 1)
    Case "T"
        tcMain(1).Visible = True
    Case "F"
        txtAplFx.Locked = False
    Case "C"
End Select

End Sub

Private Sub cboTipoConcepto_Click()
On Error GoTo vError


If vPaso Then Exit Sub

If Not vEdita Or Trim(txtPrioridad.Text) = "" Then
    txtPrioridad.Text = fxPeriodad_Id_New(Mid(cboTipoConcepto.Text, 1, 1))
End If

chkAplicacionGeneral.Enabled = True
chkSalarioImbargable.Enabled = True

chkActivo.Enabled = True
chkNoVence.Enabled = True

chkCreditoFiscal.Enabled = True
chkPermiteCambios.Enabled = True

chkMuestraENColilla.Enabled = True
chkPatronoConcepto.Enabled = True

chkNoContabiliza.Enabled = True

lblCuentaCierre.Visible = False
txtCuentaCierreId.Visible = False
txtCuentaCierreDesc.Visible = False

Select Case Mid(cboTipoConcepto.Text, 1, 1)
    Case "S" 'Salario
    Case "I" 'Ingresos
    Case "E" 'Egresos
    Case "P" 'Provisión
        chkMuestraENColilla.Value = xtpUnchecked
        
        chkPatronoConcepto.Value = xtpChecked
        chkPatronoConcepto.Enabled = False
        
        lblCuentaCierre.Visible = True
        txtCuentaCierreId.Visible = True
        txtCuentaCierreDesc.Visible = True
        
    Case "C" 'Crédito Fiscal
        chkNoContabiliza.Value = xtpChecked
        chkNoContabiliza.Enabled = False
End Select

Exit Sub

vError:

End Sub


Private Sub chkNoContabiliza_Click()
'TODO
End Sub

Private Sub chkNoVence_Click()
If chkNoVence.Value = xtpChecked Then
    dtpFinaliza.Enabled = False
Else
    dtpFinaliza.Enabled = True
End If
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 COD_CONCEPTO from RH_CONCEPTOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_CONCEPTO > '" & txtCodigo.Text & "' order by COD_CONCEPTO asc"
    Else
       strSQL = strSQL & " where COD_CONCEPTO < '" & txtCodigo.Text & "' order by COD_CONCEPTO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Cod_Concepto
      Call sbConsulta(txtCodigo.Text)
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 23
End Sub


Private Sub Form_Load()
vModulo = 23
 
vEdita = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1300
    .Add , , "Descripción", 6000
End With

cboTipoConcepto.Clear
cboTipoConcepto.AddItem "Ingresos"
cboTipoConcepto.AddItem "Egresos"
cboTipoConcepto.AddItem "Provisión"
cboTipoConcepto.AddItem "Salario"
cboTipoConcepto.Text = "Ingresos"

cboAplTipo.Clear
cboAplTipo.AddItem "Porcentual"
cboAplTipo.AddItem "Monto"
cboAplTipo.AddItem "Tabla"
cboAplTipo.AddItem "Formula"
cboAplTipo.Text = "Porcentual"

cboAplBase.Clear
cboAplBase.AddItem "Salario Ordinario"
cboAplBase.ItemData(cboAplBase.ListCount - 1) = "So"

cboAplBase.AddItem "Salario - Crédito Fiscal"
cboAplBase.ItemData(cboAplBase.ListCount - 1) = "Scf"

cboAplBase.AddItem "Salario Resultante"
cboAplBase.ItemData(cboAplBase.ListCount - 1) = "Sr"

cboAplBase.AddItem "Salario Mensual"
cboAplBase.ItemData(cboAplBase.ListCount - 1) = "Sm"

cboAplBase.AddItem "Concepto Vinculado"
cboAplBase.ItemData(cboAplBase.ListCount - 1) = "Cv"

cboAplBase.Text = "Salario Ordinario"


Dim strSQL As String

strSQL = "select rtrim(COD_GRUPO) AS 'IdX',rtrim(descripcion) as 'ItmX'" _
       & " from RH_CONCEPTOS_GRUPOS where Activo = 1 order by COD_GRUPO"

Call sbCbo_Llena_New(cboGrupo, strSQL, False, True)

cboAplResume.Clear
cboAplResume.AddItem "NA"
cboAplResume.AddItem "SALARIO_ORDINARIO"


cboAplResume.AddItem "SALARIO_DEVENGADO"
cboAplResume.AddItem "SALARIO_NETO"
cboAplResume.AddItem "INGRESOS"
cboAplResume.AddItem "EGRESOS"
cboAplResume.AddItem "HRS_ORDINARIAS"
cboAplResume.AddItem "HRS_EXTRAS"
cboAplResume.AddItem "HRS_DOBLES"
cboAplResume.AddItem "HRS_PERMISOS_CGS"
cboAplResume.AddItem "HRS_PERMISOS_SGS"
cboAplResume.AddItem "HRS_VACACIONES"
cboAplResume.AddItem "HRS_INCAPACIDAD"
cboAplResume.AddItem "VACACIONES_PERIODO"
cboAplResume.AddItem "VACACIONES_DISFRUTADAS"
cboAplResume.AddItem "VACACIONES_LIQUIDADAS"
cboAplResume.AddItem "FERIADOS"

cboAplResume.Text = "NA"


Call sbToolBarIconos(tlb, False)
Call sbToolBar(tlb, "nuevo")

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Call sbLimpiaDatos

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpiaDatos()

vCodigo = ""

tcMain.Item(0).Selected = True

tcMain.Item(3).Visible = False
tcMain.Item(4).Visible = False


btnDetalle.Visible = False
btnExclusion.Visible = False

txtCodigo.Text = ""
txtDescripcion.Text = ""

txtPrioridad.Text = ""

chkActivo.Value = xtpChecked

chkAplicacionGeneral.Value = xtpChecked
chkCreditoFiscal.Value = xtpUnchecked
chkMuestraENColilla = xtpChecked
chkNoContabiliza = xtpUnchecked
chkPatronoConcepto = xtpUnchecked
chkPermiteCambios = xtpChecked

chkSalarioImbargable.Value = xtpChecked

dtpInicio.Value = Date
chkNoVence.Value = xtpChecked

cboAplBase.Text = "Salario Ordinario"
cboTipoConcepto.Text = "Ingresos"
cboAplTipo.Text = "Porcentual"

cboAplResume.Text = "NA"
  
txtAplValor.Text = Format(0, "Standard")
txtAplFx.Text = ""
'Cuentas
txtCuentaCod.Text = ""
txtCuentaDesc.Text = ""

txtCuentaCierreId.Text = ""
txtCuentaCierreDesc.Text = ""

'Entidades y Conceptos Relacionados
txtEntidadId.Text = ""
txtEntidadDesc.Text = ""

txtConceptoVinId.Text = ""
txtConceptoVinDesc.Text = ""

'Refresca Opciones
Call chkNoVence_Click
Call chkNoContabiliza_Click

Call cboAplTipo_Click
Call cboTipoConcepto_Click

End Sub


Private Function fxTabla_Id() As Long
Dim pId As Long

strSQL = "select ISNULL(max(TABLA_ID),0) + 1 AS 'ConsecId' FROM RH_CONCEPTOS_TABLA"
Call OpenRecordSet(rs, strSQL)
    pId = rs!ConsecId
rs.Close

fxTabla_Id = pId
End Function


Private Function fxTabla_Guardar() As Long
Dim pCodigoId As Long

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxTabla_Guardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If Trim(vGrid.Text) = "" Then  'Insertar
  pCodigoId = fxTabla_Id()
  vGrid.Text = CStr(pCodigoId)
  
  strSQL = "insert into RH_CONCEPTOS_TABLA(COD_CONCEPTO,TABLA_ID,MNT_INICIO,MNT_CORTE,PORCENTAJE, ACTIVO" _
            & ", REGISTRO_USUARIO, REGISTRO_FECHA) values('" & vCodigo & "'," _
         & vGrid.Text & ","
  vGrid.col = 2
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 4
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 5
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Concepto de Planilla, Tabla Id: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update RH_CONCEPTOS_TABLA set MNT_INICIO = " & CCur(vGrid.Text) & ", MNT_CORTE = "
 vGrid.col = 3
 strSQL = strSQL & CCur(vGrid.Text) & ", PORCENTAJE = "
 vGrid.col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ", ACTIVO = "
 vGrid.col = 5
 strSQL = strSQL & vGrid.Value & " where COD_CONCEPTO = '" & vCodigo & "' AND TABLA_ID = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Concepto de Planilla, Tabla Id: " & vGrid.Text)

End If

fxTabla_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function





Private Function fxDetalle_Guardar() As Long

Dim pExiste As Long


On Error GoTo vError

fxDetalle_Guardar = 0

With gDetalle

.Row = .ActiveRow
.col = 1


strSQL = "SELECT COUNT(*) AS 'EXISTE' FROM RH_CONCEPTOS_DETALLE" _
      & " WHERE COD_CONCEPTO = '" & vCodigo & "' AND COD_DETALLE = '" & .Text & "'"
Call OpenRecordSet(rs, strSQL)
 pExiste = rs!Existe
rs.Close


If pExiste = 0 Then  'Insertar
  
  strSQL = "insert into RH_CONCEPTOS_DETALLE(COD_CONCEPTO, COD_DETALLE, DESCRIPCION, BASE_MINIMA, VALOR" _
            & ", REGISTRO_USUARIO, REGISTRO_FECHA) values('" & vCodigo & "','" & .Text & "','"
  .col = 2
  strSQL = strSQL & .Text & "',"
  .col = 3
  strSQL = strSQL & CCur(.Text) & ","
  .col = 4
  strSQL = strSQL & CCur(.Text) & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  .col = 1
  Call Bitacora("Registra", "Detalle Concepto: " & vCodigo & ", ID Detalle: " & .Text)

Else 'Actualizar

 .col = 2
 strSQL = "update RH_CONCEPTOS_DETALLE set DESCRIPCION = '" & .Text & "', BASE_MINIMA = "
 
 .col = 3
 strSQL = strSQL & CCur(.Text) & ", VALOR = "
 .col = 4
 strSQL = strSQL & CCur(.Text) & " where COD_CONCEPTO = '" & vCodigo & "' AND COD_DETALLE = '"
 .col = 1
 strSQL = strSQL & .Text & "'"

 Call ConectionExecute(strSQL)

 .col = 1
 Call Bitacora("Modifica", "Detalle Concepto: " & vCodigo & ", ID Detalle: " & .Text)

End If

End With
fxDetalle_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub sbTabla_Load()

On Error GoTo vError

strSQL = "select TABLA_ID, MNT_INICIO, MNT_CORTE, PORCENTAJE, ACTIVO" _
      & "   from RH_CONCEPTOS_TABLA" _
      & "  where COD_CONCEPTO = '" & vCodigo & "'" _
      & " order by TABLA_ID"
Call sbCargaGrid(vGrid, 5, strSQL)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub gDetalle_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Integer, strSQL As String

With gDetalle

If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxDetalle_Guardar
  
  If i = 0 Then Exit Sub
  .Row = .ActiveRow
  If .MaxRows <= .ActiveRow Then
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    .MaxRows = .MaxRows + 1
    .InsertRows .ActiveRow, 1
    .Row = .ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        .Row = .ActiveRow
        .col = 1
        If .Text <> "" Then
            strSQL = "delete RH_CONCEPTOS_DETALLE where COD_DETALLE = '" _
                    & .Text & "' AND COD_CONCEPTO = '" & vCodigo & "'"
            Call ConectionExecute(strSQL)
            strSQL = .Text
            .col = 1
            Call Bitacora("Elimina", "Detalle Concepto: " & vCodigo & ", ID Detalle: " & .Text)
            
            Call sbDetalle_List
            
        End If
     End If
End If

End With

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtExcID.Text = Item.Text
txtExcDesc.Text = Item.SubItems(1)

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 1
        Call sbTabla_Load
    Case 4
        Call sbExcluye_List
End Select
End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" And vEdita = True Then Call sbConsulta(txtCodigo.Text)
End Sub


Private Sub sbConsulta(pCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

 strSQL = "select *" _
        & " from vRH_CONCEPTOS " _
        & " where COD_CONCEPTO = '" & pCodigo & "'"
 Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  tcMain.Item(0).Selected = True
  
  tcMain.Item(3).Visible = False
  tcMain.Item(4).Visible = False

  btnDetalle.Visible = True
  btnExclusion.Visible = True
  
  
  vEdita = True

  txtCodigo.Text = rs!Cod_Concepto
  vCodigo = rs!Cod_Concepto
  
  txtDescripcion.Text = rs!Descripcion
  txtPrioridad.Text = rs!Prioridad
  
  chkActivo.Value = rs!ACTIVO
  
  chkAplicacionGeneral.Value = rs!APLICACION_GENERAL
  chkCreditoFiscal.Value = rs!APLICACION_CREDITO_FISCAL
  chkMuestraENColilla = rs!COLILLA_MUESTRA
  chkNoContabiliza = IIf((rs!REGISTRO_CONTABLE = 1), 0, 1)
  chkPatronoConcepto = rs!PATRONO_CONTROL
  chkPermiteCambios = rs!APLICACION_PERMITE_CAMBIOS
  
  chkSalarioImbargable.Value = rs!REVISA_SALARIO_IMBARGABLE
  
  dtpInicio.Value = rs!APLICACION_INICIO
  
  If IsNull(rs!APLICACION_CORTE) Then
    dtpFinaliza.Value = rs!FECHA_FINALIZA_DEFAULT
    chkNoVence.Value = xtpChecked
  Else
    dtpFinaliza.Value = rs!APLICACION_CORTE
    chkNoVence.Value = xtpUnchecked
  End If
  
  Call sbCboAsignaDato(cboGrupo, rs!GRUPO_DESC, True, rs!COD_GRUPO & "")
  Call sbCboAsignaDato(cboTipoConcepto, rs!Tipo_Concepto_Desc, True, rs!TIPO_CONCEPTO)
    
  Call sbCboAsignaDato(cboAplTipo, rs!APLICACION_TIPO_DESC, True, rs!APLICACION_TIPO)
  Call sbCboAsignaDato(cboAplBase, rs!APLICACION_BASE_DESC, True, rs!APLICACION_BASE)
  
  cboAplResume.Text = Trim(rs!SUMAR_EN & "")
    
  txtAplValor.Text = Format(rs!APLICACION_VALOR, "Standard")
  txtAplFx.Text = rs!APLICACION_FORMULA
  
  'Cuentas
  txtCuentaCod.Text = rs!CTA_MASK
  txtCuentaDesc.Text = rs!CTA_DESC
  
  txtCuentaCierreId.Text = rs!CTA_CC_Mask
  txtCuentaCierreDesc.Text = rs!CTA_CC_Desc
  
  'Entidades y Conceptos Relacionados
  txtEntidadId.Text = rs!ENTIDAD_ID
  txtEntidadDesc.Text = rs!ENTIDAD_Desc
  
  txtConceptoVinId.Text = rs!CONCEPTO_BASE_ID
  txtConceptoVinDesc.Text = rs!CONCEPTO_BASE_DESC
  
  'Refresca Opciones
  Call chkNoVence_Click
  Call chkNoContabiliza_Click
  
  Call cboAplTipo_Click
  Call cboTipoConcepto_Click
Else
  MsgBox "No se encontró registro verifique...", vbInformation
  txtCodigo.Text = ""
  txtCodigo.SetFocus
  Call sbLimpiaDatos
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        Call sbLimpiaDatos
        vEdita = False
        
        txtCodigo.Text = fxConcepto_Id_New()
        txtDescripcion.SetFocus
        
       Call sbToolBar(tlb, "edicion")
       
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     Call sbToolBar(tlb, "activo")
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaDatos
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_CONCEPTO,descripcion from RH_CONCEPTOS "
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

End Select


End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pConceptoRel As String, pEntidad As String
Dim pFechaVence As String

If Trim(txtEntidadId.Text) = "" Then
    pEntidad = "Null"
Else
    pEntidad = "'" & txtEntidadId.Text & "'"
End If

If Trim(txtConceptoVinId.Text) = "" Then
    pConceptoRel = "Null"
Else
    pConceptoRel = "'" & txtConceptoVinId.Text & "'"
End If

If chkNoVence.Value = xtpChecked Then
    pFechaVence = "Null"
Else
    pFechaVence = "'" & Format(dtpFinaliza.Value, "yyyy/mm/dd") & "'"
End If

If fxExiste(txtCodigo.Text) Then
  strSQL = "update RH_CONCEPTOS set descripcion = '" & Trim(txtDescripcion.Text) & "', PRIORIDAD = '" & Trim(txtPrioridad.Text) & "'" _
         & ",ACTIVO = " & chkActivo.Value & ",TIPO_CONCEPTO = '" & Mid(cboTipoConcepto.Text, 1, 1) & "'" _
         & ",PATRONO_CONTROL = " & chkPatronoConcepto.Value & ", COLILLA_MUESTRA = " & chkMuestraENColilla.Value _
         & ",REVISA_SALARIO_IMBARGABLE = " & chkSalarioImbargable.Value & ", REGISTRO_CONTABLE = " & IIf((chkNoContabiliza.Value = xtpChecked), 0, 1) _
         & ",APLICACION_GENERAL = " & chkAplicacionGeneral.Value & ",APLICACION_PERMITE_CAMBIOS = " & chkPermiteCambios.Value _
         & ",APLICACION_CREDITO_FISCAL = " & chkCreditoFiscal.Value _
         & ",APLICACION_TIPO = '" & Mid(cboAplTipo.Text, 1, 1) & "', APLICACION_BASE = '" & cboAplBase.ItemData(cboAplBase.ListIndex) & "'" _
         & ",APLICACION_VALOR = " & CCur(txtAplValor.Text) & ", APLICACION_FORMULA = '" & Trim(txtAplFx.Text) & "'" _
         & ",APLICACION_INICIO = '" & Format(dtpInicio.Value, "yyyy/mm/dd") & "',APLICACION_CORTE = " & pFechaVence _
         & ",SUMAR_EN = '" & cboAplResume.Text & "'" _
         & ",COD_CUENTA = '" & fxgCntCuentaFormato(False, txtCuentaCod.Text, 0) _
         & "',COD_CUENTA_CIERRE = '" & fxgCntCuentaFormato(False, txtCuentaCierreId.Text, 0) & "'" _
         & ",COD_ER = " & pEntidad _
         & ",COD_CONCEPTO_BASE = " & pConceptoRel _
         & ",COD_GRUPO = '" & cboGrupo.ItemData(cboGrupo.ListIndex) & "'" _
         & " where COD_CONCEPTO = '" & vCodigo & "' "

  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Concepto de Planilla: " & vCodigo)

Else
  vCodigo = txtCodigo.Text

   strSQL = "insert into RH_CONCEPTOS(COD_CONCEPTO,descripcion,PRIORIDAD,ACTIVO,TIPO_CONCEPTO, PATRONO_CONTROL" _
          & ", COLILLA_MUESTRA, REVISA_SALARIO_IMBARGABLE, REGISTRO_CONTABLE, APLICACION_GENERAL, APLICACION_PERMITE_CAMBIOS" _
          & ", APLICACION_CREDITO_FISCAL, APLICACION_TIPO, APLICACION_BASE, APLICACION_VALOR, APLICACION_FORMULA" _
          & ", APLICACION_INICIO, APLICACION_CORTE, SUMAR_EN, COD_CUENTA, COD_CUENTA_CIERRE" _
          & ", COD_ER, COD_CONCEPTO_BASE, COD_GRUPO, REGISTRO_USUARIO, REGISTRO_FECHA)" _
          & " values('" & vCodigo & "','" & Trim(txtDescripcion.Text) & "','" & txtPrioridad.Text & "'," & chkActivo.Value _
          & ",'" & Mid(cboTipoConcepto.Text, 1, 1) & "'," & chkPatronoConcepto.Value & "," & chkMuestraENColilla.Value _
          & "," & chkSalarioImbargable.Value & "," & IIf((chkNoContabiliza.Value = xtpChecked), 0, 1) & "," & chkAplicacionGeneral.Value _
          & "," & chkPermiteCambios.Value & "," & chkCreditoFiscal.Value _
          & ",'" & Mid(cboAplTipo.Text, 1, 1) & "','" & cboAplBase.ItemData(cboAplBase.ListIndex) _
          & "'," & CCur(txtAplValor.Text) & ",'" & Trim(txtAplFx.Text) & "'" _
          & ",'" & Format(dtpInicio.Value, "yyyy/mm/dd") & "'," & pFechaVence & ",'" & cboAplResume.Text _
          & "','" & fxgCntCuentaFormato(False, txtCuentaCod.Text, 0) & "','" & fxgCntCuentaFormato(False, txtCuentaCierreId.Text, 0) _
          & "'," & pEntidad & "," & pConceptoRel & ",'" & cboGrupo.ItemData(cboGrupo.ListIndex) & "'" _
          & ",'" & glogon.Usuario & " ',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Concepto de Planilla: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Function fxValida()

fxValida = True

If Trim(txtCodigo) = "" Then fxValida = False
If Trim(txtDescripcion) = "" Then fxValida = False

End Function



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tcMain.Item(0).Selected = True
    txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Concepto Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_CONCEPTO"
   gBusquedas.Orden = "COD_CONCEPTO"
   gBusquedas.Consulta = "select COD_CONCEPTO,descripcion from RH_CONCEPTOS"
   frmBusquedas.Show vbModal
   txtCodigo.Text = gBusquedas.Resultado
   
   tcMain.Item(0).Selected = True
   txtDescripcion.SetFocus
End If

End Sub


Private Sub txtConceptoVinId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Concepto Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_CONCEPTO"
   gBusquedas.Orden = "COD_CONCEPTO"
   gBusquedas.Consulta = "select COD_CONCEPTO,descripcion from RH_CONCEPTOS"
   gBusquedas.Filtro = " AND COD_CONCEPTO <> '" & vCodigo & "'"
   frmBusquedas.Show vbModal
   txtConceptoVinId.Text = gBusquedas.Resultado
   txtConceptoVinDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCod.Text = gCuenta
   txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaCod.Text = fxgCntCuentaFormato(True, txtCuentaCod, 0)
End If

End Sub

Private Sub txtCuentaCod_LostFocus()
   txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaCod, 0))
   txtCuentaCod.Text = fxgCntCuentaFormato(True, txtCuentaCod, 0)
End Sub

'--------------
Private Sub txtCuentaCierreId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaCierreDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCierreId.Text = gCuenta
   txtCuentaCierreDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaCierreId.Text = fxgCntCuentaFormato(True, txtCuentaCierreId, 0)
End If

End Sub

Private Sub txtCuentaCierreId_LostFocus()
   txtCuentaCierreDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaCierreId, 0))
   txtCuentaCierreId.Text = fxgCntCuentaFormato(True, txtCuentaCierreId, 0)
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoConcepto.SetFocus
End Sub



Private Sub txtEntidadId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Entidad Id"
   gBusquedas.Col2Name = "Nombre"
   gBusquedas.Col3Name = "Nombre Corto"
   gBusquedas.Columna = "COD_ER"
   gBusquedas.Orden = "COD_ER"
   gBusquedas.Consulta = "select COD_ER,NOMBRE, NOMBRE_CORTO from RH_ENTIDADES_RELACIONADAS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtEntidadId.Text = gBusquedas.Resultado
   txtEntidadDesc.Text = gBusquedas.Resultado2

End If
End Sub

Private Sub txtExcID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Concepto Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_CONCEPTO"
   gBusquedas.Orden = "COD_CONCEPTO"
   gBusquedas.Consulta = "select COD_CONCEPTO,descripcion from RH_CONCEPTOS"
   gBusquedas.Filtro = " AND COD_CONCEPTO <> '" & vCodigo & "'"
   frmBusquedas.Show vbModal
   txtExcID.Text = gBusquedas.Resultado
   txtExcDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxTabla_Guardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
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

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        If IsNumeric(vGrid.Text) Then
            strSQL = "delete RH_CONCEPTOS_TABLA where TABLA_ID = " & vGrid.Text
            Call ConectionExecute(strSQL)
            strSQL = vGrid.Text
            vGrid.col = 1
            Call Bitacora("Elimina", "Concepto de Planilla, Tabla Id: " & vGrid.Text)
            
            Call sbTabla_Load
        End If
     End If
End If


End Sub

