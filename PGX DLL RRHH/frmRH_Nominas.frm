VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Nominas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Nóminas:"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17475
   LinkTopic       =   "Form10"
   ScaleHeight     =   10065
   ScaleWidth      =   17475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lswBancos 
      Height          =   1095
      Left            =   240
      TabIndex        =   65
      Top             =   8520
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   1931
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
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbNomina_Consulta 
      Height          =   5892
      Left            =   -11160
      TabIndex        =   50
      Top             =   720
      Visible         =   0   'False
      Width           =   11652
      _Version        =   1441793
      _ExtentX        =   20553
      _ExtentY        =   10393
      _StockProps     =   79
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
      Begin XtremeSuiteControls.ListView lswConsulta 
         Height          =   4452
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Width           =   11412
         _Version        =   1441793
         _ExtentX        =   20129
         _ExtentY        =   7853
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   312
         Left            =   3960
         TabIndex        =   56
         Top             =   720
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Consulta"
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
      Begin XtremeSuiteControls.DateTimePicker dtpConsultaCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   55
         Top             =   720
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.DateTimePicker dtpConsultaInicio 
         Height          =   312
         Left            =   960
         TabIndex        =   54
         Top             =   720
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.PushButton btnConsultaClose 
         Height          =   312
         Left            =   11160
         TabIndex        =   57
         Top             =   120
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
         Appearance      =   6
         Picture         =   "frmRH_Nominas.frx":0000
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Left            =   240
         TabIndex        =   53
         Top             =   720
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fechas"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   492
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   11652
         _Version        =   1441793
         _ExtentX        =   20553
         _ExtentY        =   868
         _StockProps     =   14
         Caption         =   "Consulta de Nóminas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.44
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   3840
      Top             =   6240
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3612
      Left            =   4800
      TabIndex        =   33
      Top             =   5760
      Width           =   12012
      _Version        =   1441793
      _ExtentX        =   21188
      _ExtentY        =   6371
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
      ItemCount       =   4
      SelectedItem    =   3
      Item(0).Caption =   "Conceptos del empleado"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "lswE"
      Item(0).Control(1)=   "Label1(2)"
      Item(0).Control(2)=   "txtIngresos"
      Item(0).Control(3)=   "txtEgresos"
      Item(0).Control(4)=   "txtNeto"
      Item(0).Control(5)=   "btnConceptos"
      Item(0).Control(6)=   "chkIngresos"
      Item(0).Control(7)=   "chkEgresos"
      Item(0).Control(8)=   "btnBoletaEmail"
      Item(1).Caption =   "Patrono"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswP"
      Item(2).Caption =   "Marcas"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vgMarcas"
      Item(3).Caption =   "Registro"
      Item(3).ControlCount=   3
      Item(3).Control(0)=   "btnRegistro"
      Item(3).Control(1)=   "lsw"
      Item(3).Control(2)=   "scPersonaConceptos"
      Begin XtremeSuiteControls.ListView lswP 
         Height          =   3252
         Left            =   -70000
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   10692
         _Version        =   1441793
         _ExtentX        =   18860
         _ExtentY        =   5736
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswE 
         Height          =   3252
         Left            =   -68080
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   8772
         _Version        =   1441793
         _ExtentX        =   15473
         _ExtentY        =   5736
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2895
         Left            =   0
         TabIndex        =   59
         Top             =   840
         Width           =   10815
         _Version        =   1441793
         _ExtentX        =   19071
         _ExtentY        =   5101
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBoletaEmail 
         Height          =   312
         Left            =   -69880
         TabIndex        =   58
         Top             =   3240
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Boleta Email"
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
         Appearance      =   14
      End
      Begin XtremeSuiteControls.CheckBox chkIngresos 
         Height          =   252
         Left            =   -69880
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ingresos: "
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
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnConceptos 
         Height          =   315
         Left            =   -69880
         TabIndex        =   39
         Top             =   2880
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "+/- Conceptos"
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
         Appearance      =   14
      End
      Begin XtremeSuiteControls.FlatEdit txtIngresos 
         Height          =   312
         Left            =   -69880
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEgresos 
         Height          =   312
         Left            =   -69880
         TabIndex        =   37
         Top             =   1560
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNeto 
         Height          =   312
         Left            =   -69880
         TabIndex        =   38
         Top             =   2280
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vgMarcas 
         Height          =   3252
         Left            =   -70000
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   10692
         _Version        =   524288
         _ExtentX        =   18860
         _ExtentY        =   5736
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         MaxRows         =   1
         SpreadDesigner  =   "frmRH_Nominas.frx":0716
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.CheckBox chkEgresos 
         Height          =   252
         Left            =   -69880
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Egresos: "
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
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnRegistro 
         Height          =   300
         Left            =   3840
         TabIndex        =   49
         Top             =   390
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Registrar"
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
         Picture         =   "frmRH_Nominas.frx":0D5E
         ImageAlignment  =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption scPersonaConceptos 
         Height          =   372
         Left            =   0
         TabIndex        =   60
         Top             =   360
         Width           =   10812
         _Version        =   1441793
         _ExtentX        =   19071
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Conceptos Registrados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.42
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   -69880
         TabIndex        =   35
         Top             =   2040
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Neto a Pagar:"
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
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   6360
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Actualizar"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   312
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7646
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.DateTimePicker dtpPago 
      Height          =   312
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.FlatEdit txtNominaNo 
      Height          =   552
      Left            =   1320
      TabIndex        =   8
      Top             =   720
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   974
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPagoNo 
      Height          =   552
      Left            =   3720
      TabIndex        =   10
      Top             =   720
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   974
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   312
      Left            =   240
      TabIndex        =   12
      Top             =   5280
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7641
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   912
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7641
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
   Begin XtremeSuiteControls.FlatEdit txtNotaAlPie 
      Height          =   912
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7641
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
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   372
      Index           =   1
      Left            =   1800
      TabIndex        =   19
      Top             =   6360
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Autorizar"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   372
      Index           =   2
      Left            =   3360
      TabIndex        =   20
      Top             =   6360
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Pagar"
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
      Appearance      =   6
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4455
      Left            =   4800
      TabIndex        =   21
      Top             =   720
      Width           =   12255
      _Version        =   524288
      _ExtentX        =   21616
      _ExtentY        =   7858
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   24
      SpreadDesigner  =   "frmRH_Nominas.frx":147E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   492
      Left            =   12840
      TabIndex        =   22
      Top             =   120
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmRH_Nominas.frx":263C
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   495
      Left            =   15600
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Informe"
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
      Picture         =   "frmRH_Nominas.frx":2D3C
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   492
      Left            =   14040
      TabIndex        =   24
      Top             =   120
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmRH_Nominas.frx":3443
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   6720
      TabIndex        =   25
      Top             =   240
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   8520
      TabIndex        =   26
      Top             =   240
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7429
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
   Begin XtremeSuiteControls.FlatEdit txtEmpleadoId 
      Height          =   312
      Left            =   4920
      TabIndex        =   27
      Top             =   240
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
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   31
      Top             =   7080
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Boleta de Pago Email"
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
   End
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   32
      Top             =   7080
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Boleta de Pago Imprime"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
      Height          =   312
      Left            =   1800
      TabIndex        =   43
      Top             =   5880
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
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
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalCasos 
      Height          =   312
      Left            =   240
      TabIndex        =   42
      Top             =   5880
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
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
      Text            =   "0"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   61
      Top             =   7560
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informe por Persona"
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
   End
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   62
      Top             =   7560
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informe Contable"
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
   End
   Begin XtremeSuiteControls.PushButton btnNotas 
      Height          =   312
      Left            =   3360
      TabIndex        =   63
      ToolTipText     =   "Guarda Observaciones y Notas al Pie"
      Top             =   4680
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Notas"
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
      Picture         =   "frmRH_Nominas.frx":3D14
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnNomina 
      Height          =   375
      Index           =   7
      Left            =   2400
      TabIndex        =   66
      Top             =   8040
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Boleta de Control"
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
   End
   Begin XtremeSuiteControls.PushButton btnCaso_Export 
      Height          =   345
      Left            =   4800
      TabIndex        =   67
      ToolTipText     =   "Exportar"
      Top             =   5280
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmRH_Nominas.frx":4445
      ImageAlignment  =   0
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Salidas por Banco.:"
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
      Index           =   5
      Left            =   240
      TabIndex        =   64
      Top             =   8160
      Width           =   3975
   End
   Begin XtremeShortcutBar.ShortcutCaption scEmpleado 
      Height          =   372
      Left            =   4800
      TabIndex        =   46
      Top             =   5280
      Width           =   12012
      _Version        =   1441793
      _ExtentX        =   21188
      _ExtentY        =   656
      _StockProps     =   14
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
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total a Pagar.:"
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
      Height          =   252
      Index           =   3
      Left            =   3120
      TabIndex        =   45
      Top             =   5640
      Width           =   1332
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Casos.:"
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
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   44
      Top             =   5640
      Width           =   1332
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
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   0
      Left            =   8520
      TabIndex        =   30
      Top             =   0
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
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   5
      Left            =   4920
      TabIndex        =   29
      Top             =   0
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
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   4
      Left            =   6720
      TabIndex        =   28
      Top             =   0
      Width           =   1692
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nota al Pie"
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
      Height          =   372
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   1092
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
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
      Height          =   372
      Index           =   7
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1452
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado.:"
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
      Height          =   372
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   1092
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago No.:"
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
      Height          =   372
      Index           =   4
      Left            =   2640
      TabIndex        =   11
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina No.:"
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
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago"
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
      Height          =   252
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   732
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
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
      Height          =   252
      Index           =   4
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   732
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   732
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina"
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
      Height          =   372
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmRH_Nominas.frx":45AF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4728
   End
End
Attribute VB_Name = "frmRH_Nominas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbNomina_Consulta_Personas()

If btnNomina.Item(0).Tag = "0" And btnNomina.Item(1).Tag = "0" Then
    vGrid.MaxRows = 0
    Exit Sub
End If

'Carga Listado
With vGrid
    .MaxRows = 0
    
    strSQL = "exec spRH_Nomina_Consulta_Detalle '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & txtNominaNo.Text _
            & ",'" & Mid(Trim(txtEmpleadoId.Text), 1, 30) _
            & "','" & Mid(Trim(txtIdentificacion.Text), 1, 30) _
            & "','" & Mid(Trim(txtNombre.Text), 1, 100) & "'"
     Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       
       .Col = 3 'Empleado Id
       .Text = rs!Empleado_ID
       
       .Col = 4 'Identificacion
       .Text = rs!IDENTIFICACION
       
       .Col = 5 'Nombre
       .Text = rs!NOMBRE_COMPLETO
       .Col = 6 'Salario Nómina
       .Text = Format(rs!Salario_Ordinario, "Standard")
       .Col = 7 'Salario Mensual
       .Text = Format(rs!Salario_Devengado, "Standard")
       .Col = 8 'Total Ingresos
       .Text = Format(rs!Ingresos, "Standard")
       .Col = 9 'Total Egresos
       .Text = Format(rs!Egresos, "Standard")
       .Col = 10 'Monto a Pagar
       .Text = Format(rs!Salario_Neto, "Standard")
       .Col = 11 'Cuenta Bancaria
       .Text = Trim(rs!CUENTA_BANCARIA & "")
       .Col = 12 'Patronal
       .Text = Format(rs!Patronal, "Standard")
       
       .Col = 13 'Horas Total
       .Text = Format(rs!Hrs_Ordinarias, "Standard")
       .Col = 14 'Hrs Extras Simples
       .Text = Format(rs!Hrs_Extra, "Standard")
       .Col = 15 'Hrs Extras Dobles
       .Text = Format(rs!Hrs_Dobles, "Standard")
       .Col = 16 'Horas Permisos
       .Text = Format(rs!Hrs_Permisos_CGS + rs!Hrs_Permisos_SGS, "Standard")
       .Col = 17 'Horas Incapacidades
       .Text = Format(rs!hrs_Incapacidad, "Standard")
       .Col = 18 'Horas Vacaciones
       .Text = Format(rs!Hrs_Vacaciones, "Standard")
       
       .Col = 19 'Vacaciones Acumuladas
       .Text = Format(rs!Vacaciones_Dias_Acum, "Standard")
       .Col = 20 'Vacaciones Periodo
       .Text = Format(rs!Vacaciones_Periodo, "Standard")
       .Col = 21 'Vacaciones Disfrutadas
       .Text = Format(rs!Vacaciones_Disfrutadas, "Standard")
       .Col = 22 'Vacaciones Liquidadas
       .Text = Format(rs!Vacaciones_Liquidadas, "Standard")
       
       
       .Col = 23 'Dependientes
       
       If Not IsNull(rs!Dependientes) Then
           .Text = CStr(rs!Dependientes)
       Else
           .Text = "0"
       End If
       
       .Col = 24 'Conyuge
       If Not IsNull(rs!Conyuge) Then
           If rs!Conyuge = 1 Then
               .Text = "Sí"
           Else
               .Text = "No"
           End If
       Else
          .Text = "No ind."
       End If
    
     rs.MoveNext
    Loop
    rs.Close

End With
End Sub


Private Sub sbBoleta(pEmpleado As String)

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH: Boleta de Pago"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
'    .Formulas(2) = "Detalle = '" & strDetalle & "'"
'    .Formulas(3) = "Usuario = 'Usuario..:" & glogon.Usuario & "'"
'    .Formulas(4) = "Fecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Boleta_Pago.rpt")
    strSQL = "{vRH_Nomina_Boleta_Encabezado.COD_NOMINA} = '" & cboNomina.ItemData(cboNomina.ListIndex) _
            & "' AND {vRH_Nomina_Boleta_Encabezado.NOMINA_NUM} = " & txtNominaNo.Text
                
    If pEmpleado <> "" Then
        strSQL = strSQL & " AND {vRH_Nomina_Boleta_Encabezado.EMPLEADO_ID} = '" & pEmpleado & "'"
    End If
        
        
     .SelectionFormula = strSQL
    .PrintReport
End With

End Sub



Private Sub sbNomina_Consulta()

'Datos Principales
strSQL = "exec spRH_Nomina_Consulta_Main '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & txtNominaNo.Text
Call OpenRecordSet(rs, strSQL)

txtEstado.Text = rs!EstadoDesc
txtEstado.Tag = rs!Estado
txtPagoNo.Text = rs!NPAGO_MES

txtNotas.Text = rs!notas
txtNotaAlPie.Text = rs!Boleta_Pie_Pagina



dtpInicio.Value = rs!Fecha_Inicio
dtpCorte.Value = rs!Fecha_Corte
dtpPago.Value = rs!Fecha_Pago

txtTotalCasos.Text = Format(rs!NOMINA_TOTAL_CASOS, "###,###,##0")
txtTotalPagar.Text = Format(rs!NOMINA_TOTAL_PAGOS, "Standard")

rs.Close


dtpInicio.Enabled = False
dtpCorte.Enabled = False

btnNomina.Item(0).Enabled = False 'Actualiza
btnNomina.Item(1).Enabled = False 'Autoriza
btnNomina.Item(2).Enabled = False 'Paga
btnNomina.Item(3).Enabled = False 'Boleta Email
btnNomina.Item(4).Enabled = False 'Boleta Imprime

btnNomina.Item(5).Enabled = False 'Resumen por Persona
btnNomina.Item(6).Enabled = False 'Informe Contable


btnNomina.Item(7).Enabled = True 'Boleta de Control> Siempre Visible

Select Case txtEstado.Tag
    Case "C" 'Calculada
        btnNomina.Item(0).Enabled = True 'Actualiza
        btnNomina.Item(1).Enabled = True 'Autoriza
    
    Case "A" 'Autorizada
        btnNomina.Item(2).Enabled = True 'Paga
   
        btnNomina.Item(5).Enabled = True 'Resumen por Persona
        btnNomina.Item(6).Enabled = True 'Informe Contable
    
    Case "D" 'Denegada
        btnNomina.Item(3).Enabled = True 'Boleta Email
    
    Case "P" 'Pagada
        btnNomina.Item(3).Enabled = True 'Boleta Email
        btnNomina.Item(4).Enabled = True 'Boleta Imprime
        
        btnNomina.Item(5).Enabled = True 'Resumen por Persona
        btnNomina.Item(6).Enabled = True 'Informe Contable

End Select


'Resumen de Salidas por Banco
strSQL = "exec spRH_Nomina_Pago_Banco_Rsm '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & txtNominaNo.Text
Call OpenRecordSet(rs, strSQL)

With lswBancos.ListItems
    .Clear
  Do While Not rs.EOF
    Set itmX = .Add(, , rs!Descripcion)
        itmX.SubItems(1) = Format(rs!Casos, "###,###0")
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
    rs.MoveNext
  Loop
  rs.Close
End With

btnConceptos.Enabled = btnNomina.Item(1).Enabled

'Permiso de Gestor de Nomina
If btnNomina.Item(0).Tag = "0" Then
    btnNomina.Item(3).Enabled = False 'Boleta Email
    btnNomina.Item(4).Enabled = False 'Boleta Imprime
    btnNomina.Item(5).Enabled = False 'Resumen por Persona
    btnNomina.Item(6).Enabled = False 'Informe Contable

    btnNomina.Item(7).Enabled = True 'Boleta de Control> Siempre Visible

Else
    btnNomina.Item(5).Enabled = True 'Resumen por Persona
    btnNomina.Item(6).Enabled = True 'Informe Contable
End If

'Permiso de Autorizador de Nomina
If btnNomina.Item(1).Tag = "0" Then
  If btnNomina.Item(0).Tag = "0" Then 'Es Gestor
    btnNomina.Item(3).Enabled = False 'Boleta Email
    btnNomina.Item(4).Enabled = False 'Boleta Imprime
    btnNomina.Item(5).Enabled = True 'Boleta de Control> Siempre Visible

    btnNomina.Item(5).Enabled = False 'Resumen por Persona
    btnNomina.Item(6).Enabled = False 'Informe Contable
  End If
Else
    btnNomina.Item(5).Enabled = True 'Resumen por Persona
    btnNomina.Item(6).Enabled = True 'Informe Contable
End If


Call RefrescaTags(Me)

Call sbNomina_Consulta_Personas

End Sub

Private Sub btnBoletaEmail_Click()

If btnNomina.Item(3).Enabled = True And scEmpleado.Tag <> "" Then

    Call sbRH_Boleta_Pago_Email(cboNomina.ItemData(cboNomina.ListIndex), txtNominaNo.Text, scEmpleado.Tag)

    MsgBox "Boleta de Pago Enviada por Email!", vbInformation
Else
    MsgBox "Consulte una Nómina Autorizada y a un Empleado para poder Enviar el Correo de Boleta de Pago", vbInformation
End If

End Sub

Private Sub btnBuscar_Click()
 Call sbNomina_Consulta_Personas
End Sub

Private Sub btnCaso_Export_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
    Case tcMain.Item(0).Selected
        Call Excel_Exportar_Lsw(lswE)
    Case tcMain.Item(1).Selected
        Call Excel_Exportar_Lsw(lswP)
    Case tcMain.Item(3).Selected
        Call Excel_Exportar_Lsw(lsw)
End Select
Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnConceptos_Click()
On Error GoTo vError

tcMain.Item(3).Selected = True
Call sbPersona_Conceptos_Lista

Exit Sub
vError:


End Sub

Private Sub sbPersona_Conceptos_Lista()
Dim pEmpleadoId As String

On Error GoTo vError


pEmpleadoId = scEmpleado.Tag

'(@EmpleadoId varchar(30) = Null, @Concepto varchar(10) = Null
'                    , @Nomina varchar(10) = Null, @Vencimiento datetime = Null
'                    , @Activo smallint = -1, @LineaId int = Null)
strSQL = "exec spRRHH_Persona_Conceptos_List '" & pEmpleadoId & "',Null, '" & cboNomina.ItemData(cboNomina.ListIndex) _
                & "', Null, 1, Null"
                   
lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Num_Linea)
       itmX.SubItems(1) = rs!Tipo_Concepto_Desc
       itmX.SubItems(2) = rs!Cod_Concepto
       itmX.SubItems(3) = rs!Concepto_Desc
       itmX.SubItems(4) = Format(rs!Monto, "Standard")
       itmX.SubItems(5) = rs!Tipo_Apl_Desc
       itmX.SubItems(6) = Format(rs!BASE, "Standard")
       itmX.SubItems(7) = Format(rs!Valor, "Standard")
       itmX.SubItems(8) = rs!Base_Apl_Desc
       
       itmX.SubItems(9) = IIf((rs!ACTIVO = 1), "Sí", "No")
       itmX.SubItems(10) = Format(rs!Registro_Fecha, "yyyy-mm-dd")
       itmX.SubItems(11) = Format(rs!Fecha_Vence & "", "yyyy-mm-dd")
       itmX.SubItems(12) = rs!Documento & ""
       itmX.SubItems(13) = Format(rs!HRS_REF, "Standard")
       itmX.SubItems(14) = Format(rs!DIAS_REF, "Standard")
    
   rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnConsulta_Click()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select NOMINA_NUM, FECHA_INICIO, FECHA_CORTE, FECHA_PAGO" _
       & ", CASE When Estado = 'A' then 'Autorizada'" _
       & "       When Estado = 'P' then 'Pagada' else 'Abierta' end as 'Estado_Desc'" _
       & " FROM RH_NOMINAS" _
       & " WHERE COD_NOMINA = '" & cboNomina.ItemData(cboNomina.ListIndex) & "'" _
       & " AND FECHA_INICIO BETWEEN '" & Format(dtpConsultaInicio.Value, "yyyy/mm/dd") _
       & " 00:00:0' AND '" & Format(dtpConsultaCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & " ORDER BY FECHA_INICIO DESC"

With lswConsulta.ListItems
    .Clear
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
    
        Set itmX = .Add(, , rs!NOMINA_NUM)
            itmX.SubItems(1) = rs!Estado_Desc
            itmX.SubItems(2) = Format(rs!Fecha_Inicio, "yyyy-mm-dd")
            itmX.SubItems(3) = Format(rs!Fecha_Corte, "yyyy-mm-dd")
            itmX.SubItems(4) = Format(rs!Fecha_Pago, "yyyy-mm-dd")
    
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

Private Sub btnConsultaClose_Click()
gbNomina_Consulta.Visible = False
End Sub

Private Sub btnExportar_Click()
 Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols
    
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "..."
    
    vHeaders.Headers(3) = "Empleado Id"
    vHeaders.Headers(4) = "Identificación"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Salario Nómina"
    vHeaders.Headers(7) = "Salario Mensual"
    vHeaders.Headers(8) = "Total Ingresos"
    vHeaders.Headers(9) = "Total Egresos"
    vHeaders.Headers(10) = "Monto a Pagar"
    vHeaders.Headers(11) = "Cuenta Bancaria"
    vHeaders.Headers(12) = "Patronal"

    vHeaders.Headers(13) = "Horas Ordinarias"
    vHeaders.Headers(14) = "Horas Extras Simples"
    vHeaders.Headers(15) = "Horas Extras Dobles"
    vHeaders.Headers(16) = "Horas Permisos"
    vHeaders.Headers(17) = "Horas Incapacidades"
    vHeaders.Headers(18) = "Horas Vacaciones"
    vHeaders.Headers(19) = "Vacaciones Acumuladas"
    vHeaders.Headers(20) = "Vacaciones Periodo"
    vHeaders.Headers(21) = "Vacaciones Disfrutadas"
    vHeaders.Headers(22) = "Vacaciones Liquidadas"

    vHeaders.Headers(23) = "No. Dependientes"
    vHeaders.Headers(24) = "Cónyuge"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Nomina_Resumen")
       
End Sub

Private Sub sbInforme_Persona()
Dim pTitulo As String


pTitulo = "Nómina: " & cboNomina.Text & "   No.: " & txtNominaNo.Text & "   Estado: " & txtEstado.Text

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH, Nómina: Informe por Presona"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fxSubTitulo = '" & pTitulo & "'"
    .Formulas(3) = "fxUsuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(4) = "fxFecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Nomina_Persona_Rsm.rpt")
    strSQL = "{vRH_Nomina_Boleta_Encabezado.COD_NOMINA} = '" & cboNomina.ItemData(cboNomina.ListIndex) _
            & "' AND {vRH_Nomina_Boleta_Encabezado.NOMINA_NUM} = " & txtNominaNo.Text
        
    .SelectionFormula = strSQL
    
    .PrintReport
    
End With



End Sub


Private Sub sbInforme_Contable()
Dim pTitulo As String


pTitulo = "Nómina: " & cboNomina.Text & "   No.: " & txtNominaNo.Text & "   Estado: " & txtEstado.Text

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH, Nómina: Informe Contable"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fxSubTitulo = '" & pTitulo & "'"
    .Formulas(3) = "fxUsuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(4) = "fxFecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Nomina_Informe_Contable.rpt")
    strSQL = "{vRH_Nomina_Informe_Contable.COD_NOMINA} = '" & cboNomina.ItemData(cboNomina.ListIndex) _
            & "' AND {vRH_Nomina_Informe_Contable.NOMINA_NUM} = " & txtNominaNo.Text
        
    .SelectionFormula = strSQL
    .PrintReport
End With


End Sub


Private Sub sbBoleta_Control()
Dim pTitulo As String

On Error GoTo vError

pTitulo = "Nómina: " & cboNomina.Text & "   No.: " & txtNominaNo.Text & "   Estado: " & txtEstado.Text

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH, Nómina: Boleta de Control"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fxSubTitulo = '" & pTitulo & "'"
    .Formulas(3) = "fxUsuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(4) = "fxFecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Nomina_Boleta_Control.rpt")
    strSQL = "{vRH_Nomina_Estado_Rsm.COD_NOMINA} = '" & cboNomina.ItemData(cboNomina.ListIndex) _
            & "' AND {vRH_Nomina_Estado_Rsm.NOMINA_NUM} = " & txtNominaNo.Text
        
    .SelectionFormula = strSQL
    .SubreportToChange = "sbBancosResumen"
    
    .StoredProcParam(0) = cboNomina.ItemData(cboNomina.ListIndex)
    .StoredProcParam(1) = txtNominaNo.Text



    .Action = 1
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnNomina_Click(Index As Integer)
Dim i As Integer, pDetalle As String


On Error GoTo vError

pDetalle = "Nómina: " & cboNomina.ItemData(cboNomina.ListIndex) & ", Id: " & txtNominaNo.Text

Select Case Index
  Case 0 'Actualizar
    Me.MousePointer = vbHourglass
    
        strSQL = "exec spRH_Nomina_Actualiza '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & txtNominaNo.Text & ",'" & glogon.Usuario & "',1"
        Call ConectionExecute(strSQL)
    
        Call Bitacora("Actualiza", pDetalle)
    
    Me.MousePointer = vbDefault
    MsgBox "Nómina Recalculada!", vbInformation
    
    Call cboNomina_Click
    
  Case 1 'Autorizar
  
    If vGrid.MaxRows = 0 Then
        Me.MousePointer = vbDefault
        MsgBox "No a cargado la información para esta Nómina, verifique!", vbExclamation
        Exit Sub
    End If
  
    i = MsgBox("Esta seguro que desea Autorizar esta Nómina: " & txtNominaNo.Text & " ?", vbYesNo)
    If i = vbYes Then
        
        Me.MousePointer = vbHourglass
               
        strSQL = "exec spRH_Nomina_Autoriza '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & txtNominaNo.Text & ",'" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Autorización de " & pDetalle)
        
        Me.MousePointer = vbDefault
        MsgBox "Esta Nómina ha sido Autorizada!", vbInformation
        
        Call cboNomina_Click
    
    End If
  
  Case 2 'Pagar
    
    i = MsgBox("Esta seguro que desea PAGAR esta Nómina: " & txtNominaNo.Text & " ?", vbYesNo)
    If i = vbYes Then
    
        Me.MousePointer = vbHourglass
        
        strSQL = "exec spRH_Nomina_Pago '" & cboNomina.ItemData(cboNomina.ListIndex) & "'," & txtNominaNo.Text & ",'" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Pago de " & pDetalle)
        
        Me.MousePointer = vbDefault
        MsgBox "Nómina ha sido Reportada a Bancos para su pago!", vbInformation
        
        Call cboNomina_Click
    End If
  
  Case 3 'Boleta Envio por Email
    
    Call sbRH_Boleta_Pago_Email(cboNomina.ItemData(cboNomina.ListIndex), txtNominaNo.Text, "")
    
    Call Bitacora("Aplica", "Envío por Email de " & pDetalle)
    
    Me.MousePointer = vbDefault
    MsgBox "Boletas de Pago, fueron enviadas por Email a los Empleados!", vbInformation
  
  Case 4 'Boleta Impresora
  
    Call sbBoleta("")
  
  Case 5 'Informe por Persona
    
    Call sbInforme_Persona
    
  Case 6 'Informe Contable
    
    Call sbInforme_Contable
  
  Case 7 'Boleta de Control
    
    Call sbBoleta_Control
  

End Select



Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub



Private Sub btnNotas_Click()

On Error GoTo vError

strSQL = "Update RH_NOMINAS SET BOLETA_PIE_PAGINA = '" & txtNotaAlPie.Text _
       & "', NOTAS = '" & txtNotas.Text & "'" _
       & " Where COD_NOMINA = '" & cboNomina.ItemData(cboNomina.ListIndex) _
       & "' and NOMINA_NUM = " & txtNominaNo.Text

Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Observaciones y Notas al Pie para Nómina.: " _
        & cboNomina.ItemData(cboNomina.ListIndex) & ", Id: " & txtNominaNo.Text)

MsgBox "Observaciones y Notas al Pie para Boletas de Pago actualizadas!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRegistro_Click()


GLOBALES.gTag = scEmpleado.Tag
frmRH_Persona_Conceptos_Fijos.Show vbModal



If GLOBALES.gTag3 = "1" Then
        Me.MousePointer = vbHourglass
        
        strSQL = "exec spRH_Nomina_Actualiza '" & cboNomina.ItemData(cboNomina.ListIndex) _
             & "'," & txtNominaNo.Text & ",'" & glogon.Usuario & "',1,'" & scEmpleado.Tag & "'"
        Call ConectionExecute(strSQL)
        
        Me.MousePointer = vbDefault
        MsgBox "Nómina Recalculada!", vbInformation
        
        Call sbNomina_Consulta
        
        Call sbConsulta_Conceptos(scEmpleado.Tag, 0)
        
Else
        Call sbPersona_Conceptos_Lista

End If
        
        
End Sub

Private Sub cboNomina_Click()
If vPaso Or cboNomina.ListCount = 0 Then Exit Sub

strSQL = "exec spRH_Nomina_Actual '" & cboNomina.ItemData(cboNomina.ListIndex) & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

txtNominaNo.Text = rs!NOMINA_ID

rs.Close

Call sbNomina_Consulta

End Sub

Private Sub chkEgresos_Click()
    Call sbConsulta_Conceptos(scEmpleado.Tag, 0)
End Sub

Private Sub chkIngresos_Click()
    Call sbConsulta_Conceptos(scEmpleado.Tag, 0)
End Sub

Private Sub Form_Load()

'Nomina
vPaso = True
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)
vPaso = False

tcMain.Item(0).Selected = True

With lswConsulta.ColumnHeaders
    .Clear
    .Add , , "Id Nómina", 1800
    .Add , , "Estado", 1500, vbCenter
    .Add , , "Inicio", 1500, vbCenter
    .Add , , "Corte", 1500, vbCenter
    .Add , , "Pago", 1500, vbCenter
    
End With


With lswBancos.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Casos", 800, vbRightJustify
    .Add , , "Monto", 1800, vbRightJustify
End With
lswBancos.BackColor = RGB(214, 234, 248)

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "I/E", 1200, vbCenter
    .Add , , "Código", 800
    .Add , , "Descripción", 3500
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Tipo", 1500, vbCenter
    .Add , , "Base", 1200, vbRightJustify
    .Add , , "Valor", 1200, vbRightJustify
    
    .Add , , "Aplicación", 1000, vbCenter
    .Add , , "Activo?", 1000, vbCenter
    .Add , , "Inicia", 1400, vbCenter
    .Add , , "Vence", 1400, vbCenter
    
    .Add , , "Doc.Ref.", 1400, vbCenter
    
    .Add , , "Horas Ref", 1200, vbRightJustify
    .Add , , "Días Ref", 1200, vbRightJustify
End With




With lswE.ColumnHeaders
    .Clear
    .Add , , "Código", 800
    .Add , , "Descripción", 3500
    .Add , , "Ingresos", 1800, vbRightJustify
    .Add , , "Egresos", 1800, vbRightJustify
    .Add , , "Tipo", 1500, vbCenter
    .Add , , "Base", 1200, vbRightJustify
    .Add , , "Valor", 1200, vbRightJustify
    .Add , , "Horas Ref", 1200, vbRightJustify
    .Add , , "Días Ref", 1200, vbRightJustify
End With


With lswP.ColumnHeaders
    .Clear
    .Add , , "Código", 800
    .Add , , "Descripción", 3500
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Tipo", 1500, vbCenter
    .Add , , "Base", 1200, vbRightJustify
    .Add , , "Valor", 1200, vbRightJustify
    .Add , , "Horas Ref", 1200, vbRightJustify
    .Add , , "Días Ref", 1200, vbRightJustify
End With


txtEstado.Text = "Abierta"
txtNominaNo.Text = "1"
txtPagoNo.Text = "1"

txtNotas.Text = ""
txtNotaAlPie.Text = ""

dtpInicio.Value = Date
dtpCorte.Value = dtpInicio.Value
dtpPago.Value = dtpInicio.Value

dtpConsultaCorte.Value = dtpInicio.Value
dtpConsultaInicio.Value = DateAdd("m", -36, dtpInicio.Value)

gbNomina_Consulta.Visible = False

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub Form_Resize()

On Error Resume Next

imgBanner.Height = Me.Height
vGrid.Width = Me.Width - (vGrid.Left + 250)
vGrid.Height = Me.Height - (vGrid.Top + tcMain.Height + scEmpleado.Height + 750)

scEmpleado.Top = vGrid.Top + vGrid.Height + 100
scEmpleado.Width = vGrid.Width

btnCaso_Export.Top = scEmpleado.Top + 10

tcMain.Top = scEmpleado.Top + scEmpleado.Height + 100
tcMain.Width = vGrid.Width

lswE.Width = tcMain.Width - (lswE.Left + 100)
lswP.Width = tcMain.Width - 100
vgMarcas.Width = tcMain.Width - 100

lsw.Width = tcMain.Width - (lsw.Left + 100)
scPersonaConceptos.Width = lsw.Width

lswBancos.Height = Me.Height - (lswBancos.Top + 540)

End Sub




Private Sub lswConsulta_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

txtNominaNo.Text = Item.Text

Call sbNomina_Consulta

Call btnConsultaClose_Click

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Empleado
        Call sbConsulta_Conceptos(scEmpleado.Tag, Item.Index)
    Case 1 'Patronales
        Call sbConsulta_Conceptos(scEmpleado.Tag, Item.Index)
    Case 2 'Marcas
    Case 3 'Registro
        tcMain.Item(3).Visible = True
        Call sbPersona_Conceptos_Lista
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call cboNomina_Click

End Sub

Private Sub sbConsulta_Conceptos(pEmpleadoId As String, Optional pTab As Integer = 0)
Dim pFiltro As String


On Error GoTo vError

If txtNominaNo.Text = "" Or pEmpleadoId = "" Then Exit Sub

Me.MousePointer = vbHourglass

'tcMain.Item(0).Visible = True
'tcMain.Item(1).Visible = True
'tcMain.Item(2).Visible = True
'tcMain.Item(3).Visible = False

pFiltro = ""

tcMain.Item(pTab).Selected = True

If pTab = 0 Then
   
    lswE.ListItems.Clear
   
   If chkIngresos.Value = xtpChecked Then
        pFiltro = "'I','S'"
   End If
  
   If chkEgresos.Value = xtpChecked Then
        If pFiltro = "" Then
            pFiltro = pFiltro & "'E'"
        Else
            pFiltro = pFiltro & ",'E'"
        End If
   End If
  
   If pFiltro = "" Then
     Me.MousePointer = vbDefault
     Exit Sub
   End If
   
    strSQL = "select * " _
           & " From vRH_Nomina_Detalle_Conceptos" _
           & " where COD_NOMINA = '" & cboNomina.ItemData(cboNomina.ListIndex) & "'" _
           & " and NOMINA_NUM = " & txtNominaNo.Text _
           & " and EMPLEADO_ID = '" & pEmpleadoId & "'" _
           & " and TIPO_CONCEPTO in(" & pFiltro & ") AND IND_PATRONAL = 0 order by Prioridad"
    Call OpenRecordSet(rs, strSQL)
           
    
    Do While Not rs.EOF
      Set itmX = lswE.ListItems.Add(, , rs!Cod_Concepto)
          itmX.SubItems(1) = rs!Concepto_Desc
          itmX.SubItems(2) = Format(rs!Ingresos + rs!Salario_Ordinario, "Standard")
          itmX.SubItems(3) = Format(rs!Egresos, "Standard")
          itmX.SubItems(4) = rs!TIPO_DESC
          itmX.SubItems(5) = Format(rs!BASE, "Standard")
          itmX.SubItems(6) = rs!Valor
          itmX.SubItems(7) = rs!HRS_REF & ""
          itmX.SubItems(8) = rs!DIAS_REF & ""

    
      rs.MoveNext
    Loop
    rs.Close
    
End If



If pTab = 1 Then
   
    pFiltro = "'P'"
  
    strSQL = "select * " _
           & " From vRH_Nomina_Detalle_Conceptos" _
           & " where COD_NOMINA = '" & cboNomina.ItemData(cboNomina.ListIndex) & "'" _
           & " and NOMINA_NUM = " & txtNominaNo.Text _
           & " and EMPLEADO_ID = '" & pEmpleadoId & "'" _
           & " and (TIPO_CONCEPTO in(" & pFiltro & ") or IND_PATRONAL = 1) order by Prioridad"
    Call OpenRecordSet(rs, strSQL)
           
    lswP.ListItems.Clear
    
    Do While Not rs.EOF
      Set itmX = lswP.ListItems.Add(, , rs!Cod_Concepto)
          itmX.SubItems(1) = rs!Concepto_Desc
          itmX.SubItems(2) = Format(rs!PROVISION, "Standard")
          itmX.SubItems(3) = rs!TIPO_DESC
          itmX.SubItems(4) = Format(rs!BASE, "Standard")
          itmX.SubItems(5) = rs!Valor
          itmX.SubItems(6) = rs!HRS_REF & ""
          itmX.SubItems(7) = rs!DIAS_REF & ""
      rs.MoveNext
    Loop
    rs.Close
    
End If

  Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub txtEmpleadoId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbNomina_Consulta_Personas
End If

End Sub


Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbNomina_Consulta_Personas
End If

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbNomina_Consulta_Personas
End If

End Sub

Private Sub txtNominaNo_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyF4 Then
    gbNomina_Consulta.Top = txtNominaNo.Top
    gbNomina_Consulta.Left = cboNomina.Left
    gbNomina_Consulta.Visible = True
    Call btnConsulta_Click

End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 3

scEmpleado.Tag = vGrid.Text

vGrid.Col = 5
scEmpleado.Caption = vGrid.Text

vGrid.Col = 6
txtIngresos.Text = CCur(vGrid.Text)

vGrid.Col = 8
txtIngresos.Text = Format(CCur(txtIngresos.Text) + CCur(vGrid.Text), "Standard")

vGrid.Col = 9
txtEgresos.Text = vGrid.Text

vGrid.Col = 10
txtNeto.Text = vGrid.Text

Call sbConsulta_Conceptos(scEmpleado.Tag, 0)


If Col = 2 Then
        Call sbBoleta(scEmpleado.Tag)
End If


End Sub
