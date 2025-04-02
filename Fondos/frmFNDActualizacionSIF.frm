VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmFNDActualizacionSIF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualiza Operaciones de Cobro en el Sistema de Crédito"
   ClientHeight    =   2976
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   9864
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2976
   ScaleWidth      =   9864
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   732
      Left            =   7560
      TabIndex        =   2
      Top             =   1560
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Actualiza Cobro de Ahorros"
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
      Appearance      =   16
      Picture         =   "frmFNDActualizacionSIF.frx":0000
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   108
      Left            =   0
      TabIndex        =   0
      Top             =   2868
      Width           =   9864
      _ExtentX        =   17399
      _ExtentY        =   191
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sincroniza los Contratos de Ahorros con sus contraparte de Recaudaciones"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   7212
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmFNDActualizacionSIF"
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

lbl.Caption = "Sincronizando Contratos con Operaciones de Retención [Espere!]"
lbl.Refresh

strSQL = "exec spFndSincronizaContratos"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Sincronización de Fondos con Retenciones")

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

Private Sub Form_Load()

vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

lbl.Caption = "Actualiza Operaciones de Retención que se encuentran al cobro en el " _
            & "sistema de cuentas corrientes, actualizando la cuota al cobro o en su " _
            & "efecto la cancelación de la misma producto de una liquidación..."
End Sub
