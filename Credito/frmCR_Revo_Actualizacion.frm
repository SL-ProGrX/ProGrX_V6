VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_Revo_Actualizacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Créditos Revolutivos (Actualiza / Sincroniza)"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   1590
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "&Actualizar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Revo_Actualizacion.frx":0000
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmCR_Revo_Actualizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualiza_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

lbl.Alignment = 0

lbl.Caption = "Actualizando Contratos Revolutivos [Espere!]"
lbl.Refresh

'strSQL = "exec spCrdRevo_ContratosActualiza"
'Call ConectionExecute(strSQL)

'Call Bitacora("Aplica", "Actualización de Contratos Revolutivos")

lbl.Caption = "Actualización Concluida Satisfactoriamente...."
prgBar.Value = 1
prgBar.Max = 1000000

Me.MousePointer = vbDefault

Exit Sub


vError:
  lbl.Caption = "Error...."
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 3

End Sub

Private Sub Form_Load()

vModulo = 3

lbl.Caption = "Actualiza disponibles, vencimientos, ahorros enlazados y datos relacionados con el crédito revolutivo" _
            & " asociado a cada contrato vigente."
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

