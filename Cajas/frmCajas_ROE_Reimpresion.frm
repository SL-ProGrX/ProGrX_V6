VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCajas_ROE_Reimpresion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ROE: Reimpresión"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   14940
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   14895
      _Version        =   1441793
      _ExtentX        =   26273
      _ExtentY        =   9551
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10393
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3619
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
   Begin XtremeSuiteControls.CheckBox chkFecha 
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   330
      Index           =   0
      Left            =   8160
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
      _Version        =   1441793
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
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   330
      Index           =   1
      Left            =   9480
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
      _Version        =   1441793
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   11040
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmCajas_ROE_Reimpresion.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   615
      Left            =   12360
      TabIndex        =   9
      ToolTipText     =   "Exportar a Excel"
      Top             =   1080
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_ROE_Reimpresion.frx":0700
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   11040
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   233
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton btnImprimir 
      Height          =   615
      Left            =   13200
      TabIndex        =   12
      ToolTipText     =   "Exportar a Excel"
      Top             =   1080
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Imprimir"
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
      Picture         =   "frmCajas_ROE_Reimpresion.frx":086A
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta y Reimpresión de ROE's"
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
      Index           =   1
      Left            =   2160
      TabIndex        =   11
      Top             =   240
      Width           =   5505
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "frmCajas_ROE_Reimpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnImprimir_Click()
 Call sbFormsCall("frmCajas_ROE", vbModal, , , , Me, True)
End Sub

Private Sub chkFecha_Click()
If chkFecha.Value = xtpChecked Then
    dtpFecha(0).Enabled = True
Else
    dtpFecha(0).Enabled = False
End If

dtpFecha(1).Enabled = dtpFecha(0).Enabled
End Sub

Private Sub Form_Load()

vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1500, vbCenter
    .Add , , "Tipo ROE", 1500, vbCenter
    .Add , , "Cliente Id", 2100, vbCenter
    .Add , , "Depositante Id", 2100, vbCenter
    .Add , , "Nombre Depositante", 3500
    .Add , , "Fecha", 1500, vbCenter
    .Add , , "Usuario", 2500
    .Add , , "Monto Local", 2100, vbRightJustify
    .Add , , "Monto Dólares", 2100, vbRightJustify
    .Add , , "Tipo Cambio", 1100, vbRightJustify


End With


Call chkFecha_Click

End Sub
