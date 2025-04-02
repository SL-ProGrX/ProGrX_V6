VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmSIF_ActualizaDatosCtaCor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualización de Datos"
   ClientHeight    =   3504
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3504
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   168
      Left            =   0
      TabIndex        =   0
      Top             =   3336
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdAplica 
      Height          =   852
      Left            =   7200
      TabIndex        =   2
      Top             =   1920
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Aplicar Revisión"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmSIF_ActualizaDatosCtaCor.frx":0000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revisión y Actualización de Diferencias menores"
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
      Height          =   480
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   7212
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualiza Movimientos Ligados entre Módulos, Saldos con decimales, Traslados de Cuentas y Otras Actualizaciones..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   852
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmSIF_ActualizaDatosCtaCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplica_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass

lbl.Caption = "Actualizando datos y corrigiendo inconsistencias menores..."
DoEvents

strSQL = "exec spCRDActualizaDatos"
Call ConectionExecute(strSQL)

lbl.Caption = "Proceso Finalizado..."

Me.MousePointer = vbDefault

MsgBox "Proceso Terminado Satisfactoriamente...", vbInformation

End Sub

Private Sub Form_Load()

Dim strSQL As String

vModulo = 10
 Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub
