VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmUS_Access_Horarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Accesos: Horarios de Trabajo"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   0
      Left            =   6240
      TabIndex        =   28
      Top             =   5760
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Nuevo"
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
      Picture         =   "frmUS_Access_Horarios.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.GroupBox gbHorario 
      Height          =   3135
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   8655
      _Version        =   1441793
      _ExtentX        =   15266
      _ExtentY        =   5530
      _StockProps     =   79
      Caption         =   "Horario de Trabajo:"
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
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   0
         Left            =   2160
         TabIndex        =   8
         Top             =   480
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   1
         Left            =   3360
         TabIndex        =   9
         Top             =   480
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   2
         Left            =   2160
         TabIndex        =   11
         Top             =   840
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   3
         Left            =   3360
         TabIndex        =   12
         Top             =   840
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   4
         Left            =   2160
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   5
         Left            =   3360
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   6
         Left            =   2160
         TabIndex        =   17
         Top             =   1560
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   7
         Left            =   3360
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   8
         Left            =   2160
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   9
         Left            =   3360
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   10
         Left            =   2160
         TabIndex        =   23
         Top             =   2280
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   11
         Left            =   3360
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   12
         Left            =   2160
         TabIndex        =   26
         Top             =   2640
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHorario 
         Height          =   330
         Index           =   13
         Left            =   3360
         TabIndex        =   27
         Top             =   2640
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Format          =   2
         UpDown          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtR_Usuario 
         Height          =   330
         Left            =   6480
         TabIndex        =   32
         Top             =   840
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtR_Fecha 
         Height          =   330
         Left            =   6480
         TabIndex        =   33
         Top             =   1200
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtM_Usuario 
         Height          =   330
         Left            =   6480
         TabIndex        =   36
         Top             =   2280
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtM_Fecha 
         Height          =   330
         Left            =   6480
         TabIndex        =   37
         Top             =   2640
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   14
         Left            =   5040
         TabIndex        =   40
         Top             =   1920
         Width           =   3615
         _Version        =   1441793
         _ExtentX        =   6376
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ultima modificación:"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   5400
         TabIndex        =   39
         Top             =   2280
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuairo"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   12
         Left            =   5400
         TabIndex        =   38
         Top             =   2640
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   11
         Left            =   5400
         TabIndex        =   35
         Top             =   840
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuairo"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   5400
         TabIndex        =   34
         Top             =   1200
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   31
         Top             =   480
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registro Inicial:"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   25
         Top             =   2640
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Domingo"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   22
         Top             =   2280
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sábado"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   19
         Top             =   1920
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Viernes"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   16
         Top             =   1560
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Jueves"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   13
         Top             =   1200
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Miércoles"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Martes"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Lunes"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Activo?"
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   1
      Left            =   7560
      TabIndex        =   29
      Top             =   5760
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
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
      Picture         =   "frmUS_Access_Horarios.frx":0720
      ImageAlignment  =   4
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   2760
      TabIndex        =   30
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Descripción:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Horario Id:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Definición de Horarios"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11772
   End
End
Attribute VB_Name = "frmUS_Access_Horarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbInicializa()
Dim i As Integer

txtCodigo.Text = ""
txtDescripcion.Text = ""
chkActivo.Value = xtpChecked

For i = 0 To dtpHorario.Count - 1
    dtpHorario(i).Value = "06:00:00"
    dtpHorario(i + 1).Value = "22:00:00"
  i = i + 1
Next i

txtR_Usuario.Text = ""
txtR_Fecha.Text = ""

txtM_Usuario.Text = ""
txtM_Fecha.Text = ""



txtCodigo.SetFocus

End Sub

Private Sub sbGuardar()
Dim i As Integer, pHorarios As String

On Error GoTo vError

strSQL = "exec spPGX_Cliente_Horarios_Registra " & gPortal.Empresa_Id & ", '" & txtCodigo.Text & "', '" _
       & txtDescripcion.Text & "', " & chkActivo.Value _

pHorarios = ""
For i = 0 To dtpHorario.Count - 1
    pHorarios = pHorarios & ", '" & dtpHorario(i).Value & "'"
Next i

strSQL = strSQL & pHorarios & ", '" & glogon.Usuario & "'"

Call ConectionExecute(strSQL, 1)

MsgBox "Horario Registrado Satisfactoriamente!", vbInformation


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
 
End Sub


Private Sub btnAccion_Click(Index As Integer)
Select Case Index
    Case 0 'Nuevo
      Call sbInicializa
    Case 1 'Guardar
      Call sbGuardar
End Select
End Sub




Private Sub FlatScrollBar_Change()

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_Horario from PGX_CLIENTES_HORARIOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_empresa = " & gPortal.Empresa_Id & " and  cod_Horario > '" & txtCodigo.Text & "' order by cod_Horario asc"
    Else
       strSQL = strSQL & " where cod_empresa = " & gPortal.Empresa_Id & " and cod_Horario < '" & txtCodigo.Text & "' order by cod_Horario desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_Horario
      Call txtCodigo_LostFocus
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()

Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub


Private Sub sbConsulta()

If txtCodigo.Text = "" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from PGX_CLIENTES_HORARIOS " _
       & " Where cod_Empresa = " & gPortal.Empresa_Id & " and cod_horario = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL, 1)

If Not rs.BOF And Not rs.EOF Then
    txtDescripcion.Text = rs!Descripcion
    chkActivo.Value = rs!Activo

    dtpHorario(0).Value = Mid(rs!L_Inicio, 1, 8)
    dtpHorario(1).Value = Mid(rs!L_Corte, 1, 8)
    
    dtpHorario(2).Value = Mid(rs!K_Inicio, 1, 8)
    dtpHorario(3).Value = Mid(rs!K_Corte, 1, 8)
    
    dtpHorario(4).Value = Mid(rs!M_Inicio, 1, 8)
    dtpHorario(5).Value = Mid(rs!M_Corte, 1, 8)
    
    dtpHorario(6).Value = Mid(rs!J_Inicio, 1, 8)
    dtpHorario(7).Value = Mid(rs!J_Corte, 1, 8)
    
    dtpHorario(8).Value = Mid(rs!V_Inicio, 1, 8)
    dtpHorario(9).Value = Mid(rs!V_Corte, 1, 8)
    
    dtpHorario(10).Value = Mid(rs!S_Inicio, 1, 8)
    dtpHorario(11).Value = Mid(rs!S_Corte, 1, 8)
    
    dtpHorario(12).Value = Mid(rs!D_Inicio, 1, 8)
    dtpHorario(13).Value = Mid(rs!D_Corte, 1, 8)
    
    txtR_Usuario.Text = rs!Registro_Usuario & ""
    txtR_Fecha.Text = rs!Registro_Fecha & ""
    
    txtM_Usuario.Text = rs!Modifica_Usuario & ""
    txtM_Fecha.Text = rs!Modifica_Fecha & ""
End If



Me.MousePointer = vbDefault

Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox Err.Description, vbCritical
  
  
End Sub

Private Sub txtCodigo_LostFocus()

Call sbConsulta
End Sub
