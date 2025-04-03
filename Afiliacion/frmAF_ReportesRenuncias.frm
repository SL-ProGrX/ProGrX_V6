VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_ReportesRenuncias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Renuncias"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
   HelpContextID   =   1011
   Icon            =   "frmAF_ReportesRenuncias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   13230
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4692
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   3852
      _Version        =   1441793
      _ExtentX        =   6794
      _ExtentY        =   8276
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
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbInforme 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   13095
      _Version        =   1441793
      _ExtentX        =   23098
      _ExtentY        =   1508
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
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   615
         Left            =   11400
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Informe"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_ReportesRenuncias.frx":000C
      End
      Begin XtremeSuiteControls.CheckBox chkResumen 
         Height          =   255
         Left            =   9360
         TabIndex        =   22
         Top             =   360
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Informe Resumen"
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   5520
      TabIndex        =   0
      Top             =   5760
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   5520
      TabIndex        =   1
      Top             =   6120
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.ComboBox cboCausa 
      Height          =   315
      Left            =   8520
      TabIndex        =   5
      Top             =   2040
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboTipoRen 
      Height          =   315
      Left            =   4200
      TabIndex        =   6
      Top             =   2040
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   3480
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.FlatEdit txtConUsuario 
      Height          =   315
      Left            =   5520
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtConEjecutivo 
      Height          =   315
      Left            =   8520
      TabIndex        =   12
      Top             =   4200
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      Top             =   2760
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboTipoFecha 
      Height          =   312
      Left            =   5520
      TabIndex        =   17
      Top             =   5400
      Width           =   1692
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
   Begin XtremeSuiteControls.ComboBox cboProvincia 
      Height          =   312
      Left            =   5520
      TabIndex        =   25
      Top             =   5040
      Width           =   1692
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
   Begin XtremeSuiteControls.ComboBox cboOficina 
      Height          =   315
      Left            =   8520
      TabIndex        =   27
      Top             =   2760
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboZona 
      Height          =   315
      Left            =   8520
      TabIndex        =   29
      Top             =   3480
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.CheckBox chkMortalidad 
      Height          =   255
      Left            =   8520
      TabIndex        =   31
      Top             =   5040
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Renuncia por Mortalidad"
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
   End
   Begin XtremeSuiteControls.CheckBox chkReingreso 
      Height          =   255
      Left            =   8520
      TabIndex        =   32
      Top             =   5400
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Aplica para Re-Ingreso Automático"
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
   End
   Begin XtremeSuiteControls.CheckBox chkVolver 
      Height          =   255
      Left            =   8520
      TabIndex        =   33
      Top             =   6120
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Esta dispuesto a volver afiliarse a futuro ?"
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
   End
   Begin XtremeSuiteControls.CheckBox chkTasaAjuste 
      Height          =   255
      Left            =   8520
      TabIndex        =   34
      Top             =   5760
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Aplicar Aumento de Tasas"
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
      Value           =   2
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   11
      Left            =   8520
      TabIndex        =   30
      Top             =   3240
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Zona:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   10
      Left            =   8520
      TabIndex        =   28
      Top             =   2520
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Oficina Tramita:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   9
      Left            =   4200
      TabIndex        =   26
      Top             =   5040
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Provincia:"
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
      Height          =   372
      Left            =   0
      TabIndex        =   24
      Top             =   1320
      Width           =   3972
      _Version        =   1441793
      _ExtentX        =   7006
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Informes disponibles:"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scInforme 
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   1320
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "(Seleccione un Informe)"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   21
      Top             =   2520
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Empresa/Institución:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   7
      Left            =   4560
      TabIndex        =   20
      Top             =   6120
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Corte:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   6
      Left            =   4560
      TabIndex        =   19
      Top             =   5760
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Inicio:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   5
      Left            =   4200
      TabIndex        =   18
      Top             =   5400
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha Base:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   8520
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ejecutivo:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   8520
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Causa:"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Renuncias"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   4332
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmAF_ReportesRenuncias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Function fxSQL() As String
Dim vSQL As String


Select Case Mid(cboTipoFecha.Text, 1, 3)
    Case "Reg"
        vSQL = " {vAFI_Renuncias.registro_Fecha}"
    Case "Ven"
        vSQL = "{vAFI_Renuncias.Vencimiento}"
    Case "Res"
        vSQL = "{vAFI_Renuncias.Resuelto_Fecha}"
End Select
vSQL = vSQL & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
       & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
If cboEstado.Text <> "TODOS" Then
      vSQL = vSQL & " and {vAFI_Renuncias.Estado} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
End If
  
If cboTipoRen.Text <> "TODAS" Then
      vSQL = vSQL & " and {vAFI_Renuncias.Tipo} = '" & Mid(cboTipoRen.Text, 1, 1) & "'"
End If
  
If Len(Trim(txtConUsuario.Text)) > 0 Then
      vSQL = vSQL & " and  {vAFI_Renuncias.registro_user} = '" & txtConUsuario.Text & "'"
End If

If Len(Trim(txtConEjecutivo.Text)) > 0 Then
      vSQL = vSQL & " and {vAFI_Renuncias.Ejecutivo_Desc}  = '" & txtConEjecutivo.Text & "'"
End If

If cboCausa.Text <> "TODOS" Then
      vSQL = vSQL & " and {vAFI_Renuncias.Id_Causa} = " & cboCausa.ItemData(cboCausa.ListIndex)
End If

If cboInstitucion.Text <> "TODOS" Then
      vSQL = vSQL & " and {vAFI_Renuncias.cod_Institucion} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If

If cboProvincia.Text <> "TODOS" Then
      vSQL = vSQL & " and {vAFI_Renuncias.Provincia} = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"
End If

If cboOficina.Text <> "TODOS" Then
      vSQL = vSQL & " and {vAFI_Renuncias.Cod_Oficina} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
End If

If cboZona.Text <> "TODOS" Then
      vSQL = vSQL & " and {vAFI_Renuncias.Cod_Zona} = '" & cboZona.ItemData(cboZona.ListIndex) & "'"
End If


'Check's

If chkMortalidad.Value = xtpChecked Then
    vSQL = vSQL & " and {vAFI_Renuncias.Mortalidad} = " & chkMortalidad.Value
End If

If chkReingreso.Value = xtpChecked Then
    vSQL = vSQL & " and {vAFI_Renuncias.Aplica_Reingreso} = " & chkReingreso.Value
End If

If chkVolver.Value = xtpChecked Then
    vSQL = vSQL & " and {vAFI_Renuncias.Volver} = " & chkVolver.Value
End If

If chkTasaAjuste.Value = xtpChecked Then
    vSQL = vSQL & " and {vAFI_Renuncias.Aumenta_Puntos} = True"
End If

If chkTasaAjuste.Value = xtpUnchecked Then
    vSQL = vSQL & " and {vAFI_Renuncias.Aumenta_Puntos} = False"
End If



fxSQL = vSQL

End Function


Private Sub btnInforme_Click()
Dim vTitulo As String

On Error GoTo vError

If scInforme.Tag = "" Then
    MsgBox "Seleccione un informe!", vbExclamation
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
  
Select Case scInforme.Tag
    Case "R01" 'General
        If chkResumen.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Rsm.rpt")
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Det.rpt")
        End If
   
    Case "R02" 'Causa
        If chkResumen.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Causa_Rsm.rpt")
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Causa_Det.rpt")
        End If

  
    Case "R03" 'Informe por Usuario Registro
        If chkResumen.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Usuario_Rsm.rpt")
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Usuario_Det.rpt")
        End If
    
    Case "R04" 'Informe por Ejecutivo
        If chkResumen.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Ejecutivo_Rsm.rpt")
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Ejecutivo_Det.rpt")
        End If
    
    Case "R05" 'Informe por Empresa
        If chkResumen.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Empresa_Rsm.rpt")
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Empresa_Det.rpt")
        End If
    
    Case "R06" 'Informe por Estado
        If chkResumen.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Estado_Rsm.rpt")
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Estado_Det.rpt")
        End If
    
    Case "R07" 'Informe por Provincias
        If chkResumen.Value = xtpChecked Then
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Provincia_Rsm.rpt")
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Personas_Renuncias_Provincia_Det.rpt")
        End If
    
End Select

If chkResumen.Value = xtpChecked Then
    vTitulo = "RESUMEN, Rango: " & Format(dtpInicio.Value, "dd/mm/yyyy") _
            & " al " & Format(dtpCorte.Value, "dd/mm/yyyy") _
            & ", Estado: " & cboEstado.Text & ", Causa: " & cboCausa.Text & ", Tipo: " & cboTipoRen.Text
Else
    vTitulo = "DETALLE, Rango: " & Format(dtpInicio.Value, "dd/mm/yyyy") _
            & " al " & Format(dtpCorte.Value, "dd/mm/yyyy") _
            & ", Estado: " & cboEstado.Text & ", Causa: " & cboCausa.Text & ", Tipo: " & cboTipoRen.Text
End If

    .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Usuario = '" & glogon.Usuario & "'"
    .Formulas(2) = "Fecha   = '" & fxFechaServidor & "'"
    .Formulas(3) = "Titulo  = '" & scInforme.Caption & "'"
    .Formulas(4) = "SubTitulo = '" & vTitulo & "'"
    

    .SelectionFormula = fxSQL
  
    .Action = 1
End With


Me.MousePointer = vbDefault

Exit Sub
vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 1

End Sub


Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

scInforme.Tag = ""
scInforme.Caption = "(Seleccione un Informe)"

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Informe:", 3800
lsw.HideColumnHeaders = True

With lsw.ListItems
    .Clear
    .Add , "R01", "Informe General de Renuncias"
    .Add , "R02", "Informe por Causa de Rencuncia"
    .Add , "R03", "Informe por Usuario Registro"
    .Add , "R04", "Informe por Ejecutivo"
    .Add , "R05", "Informe por Empresa"
    .Add , "R06", "Informe por Estado"
    .Add , "R07", "Informe por Provincias"
End With

 dtpInicio = Format(fxFechaServidor, "dd/mm/yyyy")
 dtpCorte = dtpInicio.Value
 
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

scInforme.Tag = Item.Key
scInforme.Caption = Item.Text

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub



Private Sub sbInicializa()
 
On Error GoTo vError
 
Me.MousePointer = vbHourglass

cboTipoRen.Clear
cboTipoRen.AddItem "TODAS"
cboTipoRen.AddItem "Asociación"
cboTipoRen.AddItem "Patronal"
cboTipoRen.Text = "TODAS"

cboTipoFecha.Clear
cboTipoFecha.AddItem "Registro"
cboTipoFecha.AddItem "Vencimiento"
cboTipoFecha.AddItem "Resolución"
cboTipoFecha.Text = "Registro"

cboEstado.Clear
cboEstado.AddItem "TODOS"
cboEstado.AddItem "Transito"
cboEstado.ItemData(cboEstado.ListCount - 1) = "T"
cboEstado.AddItem "Rescatada"
cboEstado.ItemData(cboEstado.ListCount - 1) = "R"
cboEstado.AddItem "Perdida"
cboEstado.ItemData(cboEstado.ListCount - 1) = "P"
cboEstado.AddItem "Vencida"
cboEstado.ItemData(cboEstado.ListCount - 1) = "V"
cboEstado.AddItem "Pendiente"
cboEstado.ItemData(cboEstado.ListCount - 1) = "E"
cboEstado.Text = "TODOS"

strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
Call sbCbo_Llena_New(cboProvincia, strSQL, True, True)

strSQL = "select id_causa as IdX, rtrim(descripcion) as itmX from causas_renuncias WHERE ACTIVO = 1"
Call sbCbo_Llena_New(cboCausa, strSQL, True, True)

strSQL = "select cod_institucion as 'IdX' , rtrim(descripcion) as 'ItmX' from Instituciones order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


strSQL = "select cod_Oficina as 'IdX' , rtrim(descripcion) as 'ItmX' from SIF_Oficinas order by descripcion"
Call sbCbo_Llena_New(cboOficina, strSQL, True, True)

strSQL = "select cod_Zona as 'IdX' , rtrim(descripcion) as 'ItmX' from AFI_Zonas order by descripcion"
Call sbCbo_Llena_New(cboZona, strSQL, True, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
