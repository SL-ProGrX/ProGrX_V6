VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmUS_Parametros 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parámetros para Contraseñas"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbTitle 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16325
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   615
         Left            =   7680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Guardar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmUS_Parametros.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Recordar las últimas "
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Recordatorio de Cambio de Contraseña de "
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Tiempo en Minutos para Desbloqueo"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   4
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Número de Intentos errados"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Largo Mínimo de la Contraseña de"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   6
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Largo Máximo de la Contraseña de"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   7
      Left            =   480
      TabIndex        =   9
      Top             =   3720
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Forzar el Uso de Capitalizables almenos de"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   8
      Left            =   480
      TabIndex        =   10
      Top             =   4080
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Forzar el Uso de Caracteres Especiales de"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Renovar contraseñas cada "
      BackColor       =   16777215
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   9
      Left            =   480
      TabIndex        =   11
      Top             =   4440
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
      Text            =   "Forzar el Uso de Números de"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   10
      Left            =   6000
      TabIndex        =   12
      Top             =   1200
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "dias "
      BackColor       =   16777215
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   11
      Left            =   6000
      TabIndex        =   13
      Top             =   1560
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "Contraseñas Utilizadas"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   12
      Left            =   6000
      TabIndex        =   14
      Top             =   1920
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "dias "
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   13
      Left            =   6000
      TabIndex        =   15
      Top             =   2280
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "minutos"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   14
      Left            =   6000
      TabIndex        =   16
      Top             =   2640
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "antes de bloqueo auto."
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   15
      Left            =   6000
      TabIndex        =   17
      Top             =   3000
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "dígitos"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   16
      Left            =   6000
      TabIndex        =   18
      Top             =   3360
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "dígitos"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   17
      Left            =   6000
      TabIndex        =   19
      Top             =   3720
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "dígitos"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   18
      Left            =   6000
      TabIndex        =   20
      Top             =   4080
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "dígitos $#@^&*()"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   330
      Index           =   19
      Left            =   6000
      TabIndex        =   21
      Top             =   4440
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Text            =   "dígitos"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRenovarDias 
      Height          =   330
      Left            =   4560
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      Text            =   "30"
      BackColor       =   16777152
      Alignment       =   2
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRecordatorioClaves 
      Height          =   330
      Left            =   4560
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      Text            =   "6"
      BackColor       =   16777152
      Alignment       =   2
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRecordatorioDias 
      Height          =   330
      Left            =   4560
      TabIndex        =   24
      Top             =   1920
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      Text            =   "3"
      BackColor       =   16777152
      Alignment       =   2
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDesBloqueo 
      Height          =   330
      Left            =   4560
      TabIndex        =   25
      Top             =   2280
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      Text            =   "15"
      BackColor       =   16777152
      Alignment       =   2
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNIntentos 
      Height          =   330
      Left            =   4560
      TabIndex        =   26
      Top             =   2640
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      Text            =   "3"
      BackColor       =   16777152
      Alignment       =   2
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCtaMin 
      Height          =   330
      Left            =   4560
      TabIndex        =   27
      Top             =   3000
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      Text            =   "6"
      BackColor       =   16777152
      Alignment       =   2
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCtaMax 
      Height          =   330
      Left            =   4560
      TabIndex        =   28
      Top             =   3360
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      Text            =   "15"
      BackColor       =   16777152
      Alignment       =   2
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCapChar 
      Height          =   330
      Left            =   4560
      TabIndex        =   29
      Top             =   3720
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSimChar 
      Height          =   330
      Left            =   4560
      TabIndex        =   30
      Top             =   4080
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNumChar 
      Height          =   330
      Left            =   4560
      TabIndex        =   31
      Top             =   4440
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
      MultiLine       =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gb2FA 
      Height          =   975
      Left            =   480
      TabIndex        =   33
      Top             =   4920
      Width           =   8535
      _Version        =   1441793
      _ExtentX        =   15055
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "2FA(Doble Factor de Autenticación"
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
      Begin XtremeSuiteControls.CheckBox chk2FA 
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activar 2FA"
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
      Begin XtremeSuiteControls.ComboBox cbo2FA 
         Height          =   330
         Left            =   6240
         TabIndex        =   35
         Top             =   360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Left            =   3840
         TabIndex        =   36
         Top             =   360
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Canal de Validación"
         BackColor       =   -2147483643
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros para Contraseñas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1200
      TabIndex        =   32
      Top             =   240
      Width           =   4452
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmUS_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub btnGuardar_Click()
Dim strSQL As String

On Error GoTo vError

    
strSQL = "update us_parametros set KEY_HISTORY = " & txtRecordatorioClaves _
       & ",KEY_REMAIN_DAYS = " & txtRecordatorioDias _
       & ",KEY_RENEW_DAY = " & txtRenovarDias _
       & ",TIME_LOCK = " & txtDesBloqueo & ",KEY_INTENTOS = " & txtNIntentos _
       & ",KEY_LENMAX = " & txtCtaMax & ",KEY_LENMIN = " & txtCtaMin _
       & ",KEY_CAPCHAR = " & txtCapChar & ",KEY_SIMCHAR = " & txtSimChar _
       & ",KEY_NUMCHAR = " & txtNumChar _
       & ",TFA_IND = " & chk2FA.Value & ", TFA_METODO = '" & cbo2FA.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Parámetros de Seguridad")

MsgBox "Parámetros de Seguridad Actualizados...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 13

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cbo2FA.AddItem "MAIL"
cbo2FA.AddItem "SMS"
cbo2FA.AddItem "APP"
cbo2FA.Text = "MAIL"

strSQL = "select isnull(Count(*),0) as Existe from us_parametros"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  strSQL = "insert us_parametros(KEY_LENMIN,KEY_LENMAX,KEY_RENEW_DAY,KEY_REMAIN_DAYS" _
         & ",KEY_HISTORY,TIME_LOCK,KEY_INTENTOS,KEY_CAPCHAR,KEY_SIMCHAR,KEY_NUMCHAR, TFA_IND, TFA_METODO)" _
         & " values(6,15,30,3,6,15,3,0,0,0, 0, 'MAIL')"
  Call ConectionExecute(strSQL)
End If
rs.Close

strSQL = "select * from us_parametros"
Call OpenRecordSet(rs, strSQL)

    txtRecordatorioClaves = rs!KEY_HISTORY
    txtRecordatorioDias = rs!KEY_REMAIN_DAYS
    txtRenovarDias = rs!KEY_RENEW_DAY
    txtDesBloqueo = rs!TIME_LOCK
    txtCtaMax = rs!KEY_LENMAX
    txtCtaMin = rs!KEY_LENMIN
    txtNIntentos = rs!KEY_INTENTOS
    
    txtCapChar = rs!key_CapChar
    txtSimChar = rs!key_SimChar
    txtNumChar = rs!key_NumChar
    
    chk2FA.Value = rs!TFA_IND
    cbo2FA.Text = rs!TFA_METODO
    
rs.Close

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

