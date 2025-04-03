VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCC_ActualizaDatos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualización de Datos"
   ClientHeight    =   3072
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   10188
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3072
   ScaleWidth      =   10188
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAplica 
      Height          =   852
      Left            =   8040
      TabIndex        =   0
      Top             =   1800
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      Picture         =   "frmCC_ActualizaDatos.frx":0000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizacion de Datos Relacionados"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   7212
   End
   Begin XtremeSuiteControls.Label lbl 
      Height          =   1332
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   6852
      _Version        =   1245187
      _ExtentX        =   12086
      _ExtentY        =   2350
      _StockProps     =   79
      Caption         =   "Actualiza Movimientos Ligados entre Módulos, Saldos con decimales, Traslados de Cuentas y Otras Actualizaciones..."
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
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCC_ActualizaDatos"
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
Call ConectionExecute(strSQL)

lbl.Caption = "Proceso Finalizado..."

Me.MousePointer = vbDefault

MsgBox "Proceso Terminado Satisfactoriamente...", vbInformation

End Sub

Private Sub Form_Load()
vModulo = 10 'Cuentas Corrientes

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub
