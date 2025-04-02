VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_Poliza_Proc_Envio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Pólizas: Generación de Archivos"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   7695
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   14055
      _Version        =   1441793
      _ExtentX        =   24791
      _ExtentY        =   13573
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
   End
   Begin XtremeSuiteControls.FlatEdit txtPoliza_Codigo 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
   Begin XtremeSuiteControls.FlatEdit txtPoliza_Desc 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   330
      Left            =   8640
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   10080
      TabIndex        =   7
      Top             =   1080
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Envio.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnGenerar 
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      ToolTipText     =   "Generar"
      Top             =   1080
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Envio.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnPrevista 
      Height          =   375
      Left            =   11040
      TabIndex        =   9
      ToolTipText     =   "Prevista"
      Top             =   1080
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Prevista"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Poliza_Proc_Envio.frx":0E19
   End
   Begin XtremeSuiteControls.FlatEdit txtRegistros 
      Height          =   315
      Left            =   13080
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   556
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registros:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   12240
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Proceso:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7320
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Póliza.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Generación de archivos"
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
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Polizas de Vivienda y Prendario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frmCR_Poliza_Proc_Envio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboProceso.AddItem "202406"
cboProceso.Text = "202406"

With lsw.ColumnHeaders
    .Add , , "PRIMER_NOMB", 1500
    .Add , , "APELLIDO_PAT", 1500
    .Add , , "APELLIDO_MAT", 1500
    .Add , , "SEXO", 1000, vbCenter
    .Add , , "FECHA_NACIMIENTO", 1500, vbCenter
    .Add , , "NUMERO_CEDULA", 1500, vbCenter
    .Add , , "MONTO_ASEGURADO", 1800, vbRightJustify
    .Add , , "NUMERO_DE_OPERACION", 1500, vbCenter
    .Add , , "TIPO_POLIZA", 1500, vbCenter
End With

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
lsw.Width = Me.Width - 200
lsw.Height = Me.Height - (lsw.Top + 450)


End Sub
