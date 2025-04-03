VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCO_ControlComGeneracion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cálculo de Comisiones"
   ClientHeight    =   2880
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   10488
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   10488
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdCalcular 
      Height          =   732
      Left            =   8040
      TabIndex        =   3
      Top             =   1440
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Cálcular"
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
      Picture         =   "frmCO_ControlComGeneracion.frx":0000
   End
   Begin MSComctlLib.ProgressBar prgBarX 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   1
      Top             =   2748
      Width           =   10488
      _ExtentX        =   18500
      _ExtentY        =   233
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cálculo de Comisiones por Gestión/Recuperación de cuentas atrasadas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   7212
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso de Actualización de Abonos a reconocer en el cálculo de comisiones vía antiguedad de cuotas recuperadas!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_ControlComGeneracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcular_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "exec spCbrComision_Actualiza"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Comisiones: Actualización de Recuperación")

Me.MousePointer = vbDefault
MsgBox "Proceso de Actualización de comisiones realizado satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

