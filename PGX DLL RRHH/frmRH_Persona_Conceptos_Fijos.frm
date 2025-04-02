VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmRH_Persona_Conceptos_Fijos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conceptos Registrados a la Persona"
   ClientHeight    =   8070
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2892
      Left            =   0
      TabIndex        =   22
      Top             =   5040
      Width           =   10812
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
   Begin XtremeSuiteControls.CheckBox chkVence 
      Height          =   252
      Left            =   9000
      TabIndex        =   31
      Top             =   3480
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Vence ?"
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
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   10320
      Top             =   1440
   End
   Begin XtremeSuiteControls.UpDown udHoras 
      Height          =   312
      Left            =   8640
      TabIndex        =   0
      Top             =   2280
      Width           =   252
      _Version        =   1441793
      _ExtentX        =   444
      _ExtentY        =   556
      _StockProps     =   64
      Appearance      =   14
      UseVisualStyle  =   0   'False
      BuddyControl    =   ""
      BuddyProperty   =   ""
   End
   Begin XtremeSuiteControls.FlatEdit txtTipo 
      Height          =   312
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   1692
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConceptoDesc 
      Height          =   312
      Left            =   3840
      TabIndex        =   2
      Top             =   1560
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
   Begin XtremeSuiteControls.FlatEdit txtConcepto 
      Height          =   312
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
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
   Begin XtremeSuiteControls.FlatEdit txtHoras 
      Height          =   312
      Left            =   7560
      TabIndex        =   4
      Top             =   2280
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
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
   Begin XtremeSuiteControls.FlatEdit txtDias 
      Height          =   312
      Left            =   7560
      TabIndex        =   5
      Top             =   2640
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
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
   Begin XtremeSuiteControls.PushButton btnRegistro 
      Height          =   492
      Index           =   1
      Left            =   8160
      TabIndex        =   6
      Top             =   3960
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   868
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
      Appearance      =   16
      Picture         =   "frmRH_Persona_Conceptos_Fijos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnRegistro 
      Height          =   492
      Index           =   0
      Left            =   6720
      TabIndex        =   7
      Top             =   3960
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
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
      Appearance      =   16
      Picture         =   "frmRH_Persona_Conceptos_Fijos.frx":07CD
   End
   Begin XtremeSuiteControls.FlatEdit txtTipoAplicacion 
      Height          =   312
      Left            =   4440
      TabIndex        =   8
      Top             =   1920
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7853
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
   Begin XtremeSuiteControls.FlatEdit txtUnidadRef 
      Height          =   312
      Left            =   2760
      TabIndex        =   9
      Top             =   3480
      Width           =   1692
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtBaseRef 
      Height          =   312
      Left            =   2760
      TabIndex        =   10
      Top             =   3120
      Width           =   1692
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtResultado 
      Height          =   312
      Left            =   2760
      TabIndex        =   11
      Top             =   2640
      Width           =   1692
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtValor 
      Height          =   312
      Left            =   2760
      TabIndex        =   12
      Top             =   2280
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.UpDown udDias 
      Height          =   312
      Left            =   8640
      TabIndex        =   13
      Top             =   2640
      Width           =   252
      _Version        =   1441793
      _ExtentX        =   444
      _ExtentY        =   556
      _StockProps     =   64
      Appearance      =   14
      UseVisualStyle  =   0   'False
      BuddyControl    =   ""
      BuddyProperty   =   ""
   End
   Begin XtremeSuiteControls.ComboBox cboBaseApl 
      Height          =   312
      Left            =   7560
      TabIndex        =   27
      Top             =   3120
      Width           =   1332
      _Version        =   1441793
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   312
      Left            =   7560
      TabIndex        =   30
      Top             =   3480
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.FlatEdit txtLineaId 
      Height          =   432
      Left            =   2760
      TabIndex        =   34
      Top             =   960
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   762
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
      Text            =   "0"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnNuevo 
      Height          =   432
      Left            =   3960
      TabIndex        =   35
      Top             =   960
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   762
      _StockProps     =   79
      Caption         =   "Nuevo"
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
      Picture         =   "frmRH_Persona_Conceptos_Fijos.frx":0FA5
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   372
      Left            =   7320
      TabIndex        =   36
      Top             =   960
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Activo?  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   16
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   312
      Left            =   2760
      TabIndex        =   37
      Top             =   3960
      Width           =   1692
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento Ref"
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
      Left            =   1320
      TabIndex        =   38
      Top             =   3960
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Linea Id"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   492
      Index           =   5
      Left            =   1320
      TabIndex        =   33
      Top             =   960
      Width           =   1692
   End
   Begin XtremeShortcutBar.ShortcutCaption scNomina 
      Height          =   372
      Left            =   0
      TabIndex        =   32
      Top             =   480
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   4
      Left            =   6120
      TabIndex        =   29
      Top             =   3480
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Aplicación"
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
      Left            =   6120
      TabIndex        =   28
      Top             =   3120
      Width           =   1692
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   372
      Left            =   0
      TabIndex        =   26
      Top             =   4560
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
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scIdentificacion 
      Height          =   372
      Left            =   5400
      TabIndex        =   25
      Top             =   480
      Width           =   5412
      _Version        =   1441793
      _ExtentX        =   9546
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scEmpleadoId 
      Height          =   372
      Left            =   1440
      TabIndex        =   24
      Top             =   480
      Width           =   3972
      _Version        =   1441793
      _ExtentX        =   7006
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scNombre 
      Height          =   492
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10812
      _Version        =   1441793
      _ExtentX        =   19071
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   1560
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Horas"
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
      Index           =   2
      Left            =   6120
      TabIndex        =   20
      Top             =   2280
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Dias"
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
      Index           =   3
      Left            =   6120
      TabIndex        =   19
      Top             =   2640
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Referencia"
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
      Index           =   6
      Left            =   1320
      TabIndex        =   18
      Top             =   3120
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado"
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
      Index           =   7
      Left            =   1320
      TabIndex        =   17
      Top             =   2640
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Index           =   8
      Left            =   1320
      TabIndex        =   16
      Top             =   1920
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Index           =   9
      Left            =   1320
      TabIndex        =   15
      Top             =   2280
      Width           =   1692
   End
   Begin VB.Label lblBaseDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   3480
      Width           =   1692
   End
End
Attribute VB_Name = "frmRH_Persona_Conceptos_Fijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vNominaCorte As Date, vFechaNew As Date

Dim strSQL As String, rs As New ADODB.Recordset



Private Sub sbPersona_Datos()

On Error GoTo vError

strSQL = "select Empleado_Id, Identificacion, Nombre_Completo, cod_nomina, dbo.Mygetdate() as FechaBase" _
      & " from rh_personas where Empleado_Id = '" & scEmpleadoId.Caption & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    scNombre.Caption = rs!NOMBRE_COMPLETO
    scIdentificacion.Caption = rs!IDENTIFICACION
    scNomina.Caption = rs!COD_NOMINA
    vFechaNew = rs!FechaBase
    
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbPersona_Concepto_Id(pLineaId As Long)

On Error GoTo vError

strSQL = "exec spRRHH_Persona_Conceptos_List '" & scEmpleadoId.Caption & "', Null, Null, Null, Null," & pLineaId

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
       
   txtLineaId.Text = pLineaId
       
   txtConcepto.Text = rs!Cod_Concepto
   txtConceptoDesc.Text = rs!Concepto_Desc
    
   txtTipo.Text = UCase(rs!Tipo_Apl_Desc)
   
   txtValor.Text = Format(rs!Valor, "Standard")
   
   txtResultado.Text = Format(0, "Standard")
   txtUnidadRef.Text = Format(0, "Standard")
   
   txtBaseRef.Tag = rs!SALARIO_BASE
   txtBaseRef.Text = Format(rs!SALARIO_BASE, "Standard")
   
   If rs!Tipo = "M" Then
        txtValor.Locked = False
        txtResultado.Text = Format(rs!Valor, "Standard")
        txtBaseRef.Text = Format(0, "Standard")
   End If
   
   
   txtTipoAplicacion.Text = rs!Tipo_Concepto_Desc
   

   txtUnidadRef.Text = Format(rs!UD_CALCULO_REF, "Standard")
   
   txtDias.Text = "0"
   txtHoras.Text = "0"
   
    
   Select Case Mid(rs!UD_CALCULO, 1, 1)
    Case "D" 'Dias
        udDias.Visible = True
        udHoras.Visible = False
        
        lblBaseDesc.Caption = "Valor en Días:"
        txtValor.Locked = True
        
    Case "H" 'Horas
        udHoras.Visible = True
        udDias.Visible = False
        lblBaseDesc.Caption = "Valor en Horas:"
        txtValor.Locked = True
    Case Else
        lblBaseDesc.Caption = "Unidad:"
        udDias.Visible = False
        udHoras.Visible = False
   End Select
   
   
   cboBaseApl.Text = rs!Base_Apl_Desc
   Call cboBaseApl_Click
    
   txtDocumento.Text = rs!Documento & ""
    
   chkActivo.Value = rs!ACTIVO
    
   If IsNull(rs!Fecha_Vence) Then
        chkVence.Value = xtpUnchecked
        dtpVence.Value = vFechaNew
   Else
        chkVence.Value = xtpChecked
        dtpVence.Value = rs!Fecha_Vence
   End If
   
   Call chkVence_Click
    
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbConceptos_Lista()
Dim pEmpleadoId As String
Dim itmX As ListViewItem

On Error GoTo vError


pEmpleadoId = scEmpleadoId.Caption


strSQL = "exec spRRHH_Persona_Conceptos_List '" & pEmpleadoId & "'"
                    
                  
                   
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



Private Sub sbConcepto_Consulta()
Dim pEmpleadoId As String

On Error GoTo vError


pEmpleadoId = scEmpleadoId.Caption

txtValor.Locked = True


strSQL = "exec spRRHH_Concepto_InfoBase_Persona '" & pEmpleadoId & "','" & txtConcepto.Text & "'"
                    
                    
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
   
   
   txtTipo.Text = rs!APLICACION_TIPO_DESC
   
   txtValor.Text = Format(rs!APLICACION_VALOR, "Standard")
   
   txtResultado.Text = Format(0, "Standard")
   txtUnidadRef.Text = Format(0, "Standard")
   
   txtBaseRef.Tag = rs!APLICACION_BASE
   txtBaseRef.Text = Format(rs!SALARIO_BASE, "Standard")
   
   If rs!APLICACION_TIPO = "M" Then
        txtValor.Locked = False
        txtResultado.Text = Format(rs!APLICACION_VALOR, "Standard")
        txtBaseRef.Text = Format(0, "Standard")
   End If
   
   
   txtTipoAplicacion.Text = rs!Tipo_Concepto_Desc
   

   txtUnidadRef.Text = Format(rs!UD_CALCULO_REF, "Standard")
   
   txtHoras.Locked = True
   
   txtDias.Text = "0"
   txtHoras.Text = "0"
   
    
   Select Case Mid(rs!UD_CALCULO, 1, 1)
    Case "D" 'Dias
        udDias.Visible = True
        udHoras.Visible = False
        
        lblBaseDesc.Caption = "Valor en Días:"
        txtValor.Locked = True
        
    Case "H" 'Horas
        txtHoras.Locked = False
        udHoras.Visible = True
        udDias.Visible = False
        lblBaseDesc.Caption = "Valor en Horas:"
        txtValor.Locked = True
    Case Else
        lblBaseDesc.Caption = "Unidad:"
        udDias.Visible = False
        udHoras.Visible = False
   
        If rs!APLICACION_TIPO = "P" Then
             txtValor.Locked = False
             txtResultado.Text = Format(rs!APLICACION_VALOR, "Standard")
             txtBaseRef.Text = Format(0, "Standard")
        End If
   
   End Select
   

   
   
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnNuevo_Click()

txtLineaId.Text = 0
txtConcepto.Text = ""
txtConceptoDesc.Text = ""
txtTipo.Text = ""
txtTipoAplicacion.Text = ""

txtValor.Text = 0

txtResultado.Text = 0
txtBaseRef.Text = 0

txtUnidadRef.Text = ""

txtHoras.Text = 0
txtDias.Text = 0

txtDocumento.Text = ""

chkActivo.Value = xtpChecked

cboBaseApl.Text = "Mensual"
cboBaseApl_Click

dtpVence.Value = vFechaNew

End Sub

Private Sub btnRegistro_Click(Index As Integer)
Dim pVence As String, pMonto As Currency


If txtConcepto.Text = "" Then Exit Sub

On Error GoTo vError


Me.MousePointer = vbHourglass

GLOBALES.gTag3 = "1"

Select Case Index
    Case 0 'Registro
            
            pVence = "Null"
            
            If chkVence.Value = xtpChecked Then
                pVence = "'" & Format(dtpVence.Value, "yyyy-mm-dd") & "'"
            End If
            
            pMonto = CCur(txtResultado.Text)
            
            If Mid(txtTipo.Text, 1, 1) = "M" Then
                pMonto = CCur(txtValor.Text)
            End If
            
            strSQL = "exec spRH_Persona_Concepto_Registra '" & scEmpleadoId.Caption _
                   & "','" & txtConcepto.Text & "','" & Mid(txtTipo.Text, 1, 1) & "'," & CCur(txtValor.Text) _
                   & "," & pVence & ", Null, '" & Mid(txtDocumento.Text, 1, 30) _
                   & "','" & glogon.Usuario & "','" & Mid(cboBaseApl.Text, 1, 1) & "'," & pMonto _
                   & "," & CCur(txtUnidadRef.Text) & "," & CCur(txtHoras.Text) & "," & CCur(txtDias.Text) _
                   & "," & txtLineaId.Text & "," & chkActivo.Value
                   
            Call ConectionExecute(strSQL)
        
    Case 1 'Elimina
      If txtLineaId.Text > 0 Then

            strSQL = "exec spRH_Persona_Concepto_Elimina '" & scEmpleadoId.Caption & "'," & txtLineaId.Text _
                   & ",'" & glogon.Usuario & "'"
            Call ConectionExecute(strSQL)

      End If
      
End Select

Me.MousePointer = vbDefault

Call btnNuevo_Click

txtConcepto.Text = ""
txtConceptoDesc.Text = ""

Call sbConceptos_Lista

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboBaseApl_Click()

Select Case Mid(cboBaseApl.Text, 1, 1)
    Case "N"
        chkVence.Value = xtpChecked
        chkVence.Enabled = False
        
    Case "M"
        chkVence.Value = xtpChecked
        chkVence.Enabled = True
    
End Select

Call chkVence_Click


End Sub

Private Sub chkVence_Click()

If chkVence.Value = xtpChecked Then
 dtpVence.Enabled = True
Else
 dtpVence.Enabled = False
End If

End Sub

Private Sub Form_Load()


scEmpleadoId.Caption = GLOBALES.gTag

GLOBALES.gTag3 = "0"

vPaso = True

    cboBaseApl.Clear
    cboBaseApl.AddItem "Nómina"
    cboBaseApl.AddItem "Mensual"
    cboBaseApl.Text = "Mensual"

vPaso = False


With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "I/E", 1200, vbCenter
    .Add , , "Código", 800
    .Add , , "Descripción", 3500
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Tipo", 1100, vbCenter
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

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Call sbPersona_Concepto_Id(Item.Text)

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbPersona_Datos
Call sbConceptos_Lista

Call cboBaseApl_Click

Call btnNuevo_Click


End Sub

Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 And txtLineaId.Text = "0" Then
    gBusquedas.Columna = "COD_CONCEPTO"
    gBusquedas.Orden = "COD_CONCEPTO"
    gBusquedas.Consulta = "select COD_CONCEPTO, DESCRIPCION FROM vRH_Nomina_Conceptos_Permitidos_RegMan"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    txtConcepto.Text = gBusquedas.Resultado
    txtConceptoDesc.Text = gBusquedas.Resultado2
    
    Call sbConcepto_Consulta
    
End If

End Sub



Private Sub txtHoras_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If Not IsNumeric(txtHoras.Text) Then Exit Sub

txtResultado.Text = Format(CCur(txtHoras.Text) * CDbl(txtUnidadRef.Text), "Standard")

vError:
End Sub

Private Sub txtValor_KeyUp(KeyCode As Integer, Shift As Integer)

If Mid(txtTipoAplicacion.Text, 1, 1) = "M" Then
    txtResultado.Text = txtValor.Text
End If
End Sub

Private Sub txtValor_LostFocus()

On Error GoTo vError

 txtValor.Text = Format(CCur(txtValor.Text), "Standard")
 
 txtResultado.Text = Format(CCur(txtResultado.Text), "Standard")
 
vError:

End Sub


Private Sub udDias_DownClick()
On Error GoTo vError

If CInt(txtDias.Text) > 0 Then
    txtDias.Text = CStr(CInt(txtDias.Text) - 1)
End If

txtResultado.Text = Format(CCur(txtDias.Text) * CDbl(txtUnidadRef.Text), "Standard")

vError:

End Sub

Private Sub udDias_UpClick()
On Error GoTo vError

txtDias.Text = CStr(CInt(txtDias.Text) - 1)

txtResultado.Text = Format(CCur(txtDias.Text) * CDbl(txtUnidadRef.Text), "Standard")

vError:

End Sub



Private Sub udHoras_DownClick()
On Error GoTo vError

If CInt(txtHoras.Text) > 0 Then
    txtHoras.Text = CStr(CInt(txtHoras.Text) - 1)
End If

txtResultado.Text = Format(CCur(txtHoras.Text) * CDbl(txtUnidadRef.Text), "Standard")

vError:

End Sub

Private Sub udHoras_UpClick()
On Error GoTo vError

txtHoras.Text = CStr(CInt(txtHoras.Text) + 1)

txtResultado.Text = Format(CCur(txtHoras.Text) * CDbl(txtUnidadRef.Text), "Standard")

vError:
End Sub
