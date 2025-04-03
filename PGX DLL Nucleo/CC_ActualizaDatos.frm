VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSIF_ActualizaDatosCtaCor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualización de Datos"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAplica 
      Caption         =   "&Aplicar"
      Height          =   1215
      Left            =   120
      Picture         =   "CC_ActualizaDatos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbl 
      Caption         =   "Actualiza Movimientos Ligados entre Módulos, Saldos con decimales, Traslados de Cuentas y Otras Actualizaciones..."
      Height          =   1215
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   3735
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
lbl.Refresh

strSQL = "exec spCRDActualizaDatos"
glogon.Conection.Execute strSQL

lbl.Caption = "Proceso Finalizado..."

Me.MousePointer = vbDefault

MsgBox "Proceso Terminado Satisfactoriamente...", vbInformation

End Sub
