VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmActivos_Justificaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Justificaciones (Ajustes)"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   10095
      _Version        =   1441793
      _ExtentX        =   17801
      _ExtentY        =   5948
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
      ItemCount       =   1
      Item(0).Caption =   "General"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "Label6(2)"
      Item(0).Control(1)=   "Label6(1)"
      Item(0).Control(2)=   "Label6(0)"
      Item(0).Control(3)=   "txtDescripcion"
      Item(0).Control(4)=   "cbo"
      Item(0).Control(5)=   "txtAsiento"
      Item(0).Control(6)=   "txtAsientoDesc"
      Item(0).Control(7)=   "txtCta04Desc"
      Item(0).Control(8)=   "txtCta03Desc"
      Item(0).Control(9)=   "txtCta02Desc"
      Item(0).Control(10)=   "txtCta01Desc"
      Item(0).Control(11)=   "txtCta01"
      Item(0).Control(12)=   "txtCta02"
      Item(0).Control(13)=   "txtCta03"
      Item(0).Control(14)=   "txtCta04"
      Item(0).Control(15)=   "lblC01"
      Item(0).Control(16)=   "lblC02"
      Item(0).Control(17)=   "lblC03"
      Item(0).Control(18)=   "lblC04"
      Begin XtremeSuiteControls.FlatEdit txtCta04Desc 
         Height          =   312
         Left            =   4320
         TabIndex        =   1
         Top             =   2880
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtCta03Desc 
         Height          =   312
         Left            =   4320
         TabIndex        =   2
         Top             =   2520
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtCta02Desc 
         Height          =   312
         Left            =   4320
         TabIndex        =   3
         Top             =   2160
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtCta01Desc 
         Height          =   312
         Left            =   4320
         TabIndex        =   4
         Top             =   1800
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Width           =   7812
         _Version        =   1441793
         _ExtentX        =   13779
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.FlatEdit txtAsientoDesc 
         Height          =   312
         Left            =   4320
         TabIndex        =   7
         Top             =   1440
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtAsiento 
         Height          =   312
         Left            =   2160
         TabIndex        =   8
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1440
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.FlatEdit txtCta01 
         Height          =   312
         Left            =   2160
         TabIndex        =   9
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1800
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.FlatEdit txtCta02 
         Height          =   312
         Left            =   2160
         TabIndex        =   10
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2160
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.FlatEdit txtCta03 
         Height          =   312
         Left            =   2160
         TabIndex        =   11
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2520
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.FlatEdit txtCta04 
         Height          =   312
         Left            =   2160
         TabIndex        =   12
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2880
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.Label lblC04 
         Height          =   252
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta...."
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
      End
      Begin XtremeSuiteControls.Label lblC03 
         Height          =   252
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta..."
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
      End
      Begin XtremeSuiteControls.Label lblC02 
         Height          =   252
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta..."
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
      End
      Begin XtremeSuiteControls.Label lblC01 
         Height          =   252
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta...."
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo Asiento"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   2160
      TabIndex        =   20
      Top             =   600
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   2160
      TabIndex        =   23
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   1095
      _Version        =   1441793
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
      Picture         =   "frmActivos_Justificaciones.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   3240
      TabIndex        =   24
      ToolTipText     =   "Editar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Justificaciones.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   3600
      TabIndex        =   25
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Justificaciones.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   4200
      TabIndex        =   26
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Justificaciones.frx":11D1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   4560
      TabIndex        =   27
      ToolTipText     =   "Deshacer"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Justificaciones.frx":1902
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   5040
      TabIndex        =   28
      ToolTipText     =   "Reporte"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Justificaciones.frx":2002
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.Label LabelX 
      Height          =   435
      Left            =   360
      TabIndex        =   21
      Top             =   600
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   762
      _StockProps     =   79
      Caption         =   "Justificación"
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
Attribute VB_Name = "frmActivos_Justificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean

Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String


Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpiaPantalla
        txtCodigo.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
        vEdita = True
        cbo.SetFocus
        
        Call sbBarra_Accion("Editar")
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
    
    Case 5 'REPORTES
   
End Select


End Sub




Private Sub cbo_Click()

lblC01.Caption = "Cuenta de Activo"
lblC02.Caption = "Cuenta Depreciacion"
lblC03.Caption = "Cuenta de Gastos"
lblC04.Caption = "Cuenta de Patrimonio"

lblC01.Visible = True
txtCta01.Visible = True
txtCta01Desc.Visible = True

lblC02.Visible = False
txtCta02.Visible = False
txtCta02Desc.Visible = False

lblC03.Visible = False
txtCta03.Visible = False
txtCta03Desc.Visible = False

lblC04.Visible = False
txtCta04.Visible = False
txtCta04Desc.Visible = False


  Select Case cbo.Text
    Case "Adiciones y Mejoras"
        lblC01.Caption = "Cuenta Transitoria"
    
    Case "Retiros (Salidas)"
        lblC01.Caption = "Cta Ingreso x Ganancia"
        lblC02.Caption = "Cta Gasto x Pérdida"
        lblC03.Caption = "Efectivo / Transitoria"
        
        lblC02.Visible = True
        txtCta02.Visible = True
        txtCta02Desc.Visible = True
        
        lblC03.Visible = True
        txtCta03.Visible = True
        txtCta03Desc.Visible = True
    
    Case "Revaluaciones"
        lblC01.Caption = "Cuenta de Patrimonio"
    
    Case "Deterioros y Desvalorizaciones"
        lblC01.Caption = "Cta Gasto x Deterioro"
        lblC02.Caption = "Cta Estimación x Deter."
       
        lblC02.Visible = True
        txtCta02.Visible = True
        txtCta02Desc.Visible = True
    
    Case "Mantenimiento"
        lblC01.Caption = "Cta Gasto x Manteni."
        lblC02.Caption = "Cta Estimación"
       
        lblC02.Visible = True
        txtCta02.Visible = True
        txtCta02Desc.Visible = True
  
  End Select

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_justificacion from Activos_justificaciones"

    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_justificacion > '" & txtCodigo.Text & "' order by cod_justificacion asc"
    Else
       strSQL = strSQL & " where cod_justificacion < '" & txtCodigo.Text & "' order by cod_justificacion desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_justificacion
      Call sbConsulta(txtCodigo.Text)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()

On Error GoTo vError
vModulo = 36
 

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
cbo.Clear
cbo.AddItem "Adiciones y Mejoras"
cbo.AddItem "Retiros (Salidas)"
cbo.AddItem "Revaluaciones"
cbo.AddItem "Deterioros y Desvalorizaciones"
cbo.AddItem "Mantenimiento"
 
 vEdita = False

 Call sbBarra_Accion("Nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCodigo = ""

cbo.Text = "Adiciones y Mejoras"

txtDescripcion = ""

txtCta01 = ""
txtCta01Desc = ""

txtCta02 = ""
txtCta02Desc = ""

txtCta03 = ""
txtCta03Desc = ""

txtCta04 = ""
txtCta04Desc = ""

txtAsiento.Text = ""
txtAsientoDesc.Text = ""
End Sub


Public Sub sbConsultaExterna(pJustificacion As String)
If pJustificacion <> "" Then
 Call sbConsulta(pJustificacion)
End If
End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from Activos_justificaciones where cod_justificacion = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
   Call sbBarra_Accion("activo")
  vEdita = True
  
  vCodigo = rs!cod_justificacion
  txtCodigo = rs!cod_justificacion
 
  txtDescripcion.Text = rs!Descripcion
  txtDescripcion.SetFocus
    
    
  txtAsiento.Text = rs!Tipo_Asiento
  txtAsientoDesc.Text = fxgCntTipoAsientoDesc(rs!Tipo_Asiento)
  
  Select Case rs!Tipo
    Case "A"
      cbo.Text = "Adiciones y Mejoras"
    Case "R"
      cbo.Text = "Retiros (Salidas)"
    Case "V"
      cbo.Text = "Revaluaciones"
    Case "D"
      cbo.Text = "Deterioros y Desvalorizaciones"
    Case "M"
      cbo.Text = "Mantenimiento"
  End Select
  Call cbo_Click
  
  If txtCta01.Visible Then
        txtCta01 = fxgCntCuentaFormato(True, rs!cod_cuenta_01, 0)
        txtCta01Desc = fxgCntCuentaDesc(rs!cod_cuenta_01)
  End If
  
  If txtCta02.Visible Then
        txtCta02 = fxgCntCuentaFormato(True, rs!cod_cuenta_02, 0)
        txtCta02Desc = fxgCntCuentaDesc(rs!cod_cuenta_02)
  End If
  
  If txtCta03.Visible Then
        txtCta03 = fxgCntCuentaFormato(True, rs!cod_cuenta_03, 0)
        txtCta03Desc = fxgCntCuentaDesc(rs!cod_cuenta_03)
  End If
  
  If txtCta03.Visible Then
        txtCta04 = fxgCntCuentaFormato(True, rs!cod_cuenta_04, 0)
        txtCta04Desc = fxgCntCuentaDesc(rs!cod_cuenta_04)
  End If
  
Else
  If vEdita Then
      MsgBox "No se encontró registro verifique...", vbInformation
  End If
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del tipo de Activo no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String

On Error GoTo vError

Select Case Mid(cbo.Text, 1, 3)
  Case "Adi"
    vTipo = "A"
  Case "Ret"
    vTipo = "R"
  Case "Rev"
    vTipo = "V"
  Case "Det"
    vTipo = "D"
  Case "Man"
    vTipo = "M"
End Select

If vEdita Then
  strSQL = "update Activos_justificaciones set descripcion = '" & txtDescripcion.Text _
         & "',tipo = '" & vTipo & "',Tipo_Asiento = '" & txtAsiento.Text _
         & "',cod_cuenta_01 = '" & fxgCntCuentaFormato(False, txtCta01, 0) _
         & "',cod_cuenta_02 = '" & fxgCntCuentaFormato(False, txtCta02, 0) _
         & "',cod_cuenta_03 = '" & fxgCntCuentaFormato(False, txtCta03, 0) _
         & "',cod_cuenta_04 = '" & fxgCntCuentaFormato(False, txtCta04, 0) _
         & "',Modifica_Fecha = getdate(), Modifica_Usuario = '" & glogon.Usuario _
         & "' where cod_justificacion = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Tipo de Justificación: " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into Activos_justificaciones(cod_justificacion,descripcion" _
          & ",tipo, Tipo_Asiento, cod_cuenta_01,cod_cuenta_02,cod_cuenta_03,cod_cuenta_04,registro_Fecha,Registro_Usuario) values('" _
          & vCodigo & "','" & txtDescripcion.Text & "','" & vTipo & "','" & txtAsiento.Text _
          & "','" & fxgCntCuentaFormato(False, txtCta01, 0) _
          & "','" & fxgCntCuentaFormato(False, txtCta02, 0) _
          & "','" & fxgCntCuentaFormato(False, txtCta03, 0) _
          & "','" & fxgCntCuentaFormato(False, txtCta04, 0) _
          & "',getdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Tipo de Justificación: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
 Call sbBarra_Accion("activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Activos_justificaciones where cod_justificacion = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Tipo de Justificacion : " & vCodigo)
  Call sbLimpiaPantalla
   Call sbBarra_Accion("nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta01.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
  gBusquedas.Filtro = " and COD_CONTABILIDAD = " & GLOBALES.gEnlace
  gBusquedas.Columna = "tipo_asiento"
  gBusquedas.Orden = "tipo_asiento"
  frmBusquedas.Show vbModal
  txtAsiento = gBusquedas.Resultado
  txtAsientoDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  cbo.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_justificacion"
  gBusquedas.Orden = "cod_justificacion"
  gBusquedas.Consulta = "select cod_justificacion as codigo,descripcion from Activos_justificaciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo.Text <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtCta01_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta01Desc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta01 = gCuenta
End If
End Sub

Private Sub txtCta01_LostFocus()
If txtCta01.Text <> "" Then
        txtCta01.Text = fxgCntCuentaFormato(True, txtCta01.Text)
        txtCta01Desc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCta01.Text, 0))
End If
End Sub

Private Sub txtCta01Desc_KeyDown(KeyCode As Integer, Shift As Integer)

If txtCta02.Visible Then
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta02.SetFocus
Else
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta01 = gCuenta
   txtCta01.SetFocus
End If
End Sub

Private Sub txtCta02_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta02Desc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta02 = gCuenta
End If
End Sub

Private Sub txtCta02_LostFocus()
If txtCta02.Text <> "" Then
    txtCta02.Text = fxgCntCuentaFormato(True, txtCta02.Text)
    txtCta02Desc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCta02.Text, 0))
End If
End Sub

Private Sub txtCta02Desc_KeyDown(KeyCode As Integer, Shift As Integer)
If txtCta03.Visible Then
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta03.SetFocus
Else
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta02 = gCuenta
   txtCta02.SetFocus
End If
End Sub

Private Sub txtCta03_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta03Desc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta03 = gCuenta
End If
End Sub

Private Sub txtCta03_LostFocus()
If txtCta03.Text <> "" Then
    txtCta03.Text = fxgCntCuentaFormato(True, txtCta03.Text)
    txtCta03Desc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCta03.Text, 0))
End If
End Sub

Private Sub txtCta03Desc_KeyDown(KeyCode As Integer, Shift As Integer)

If txtCta04.Visible Then
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta04.SetFocus
Else
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta03 = gCuenta
   txtCta03.SetFocus
End If

End Sub


Private Sub txtCta04_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta04Desc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta04 = gCuenta
End If
End Sub

Private Sub txtCta04_LostFocus()
If txtCta03.Text <> "" Then
    txtCta04.Text = fxgCntCuentaFormato(True, txtCta04.Text)
    txtCta04Desc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCta04.Text, 0))
End If
End Sub

Private Sub txtCta04Desc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCta04 = gCuenta
   txtCta04.SetFocus
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAsiento.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_justificacion as codigo,descripcion from Activos_justificaciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub



