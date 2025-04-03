VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmInvVentasEnLinea_Sinc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sincronizar Ventas en Línea"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnSincroniza 
      Height          =   615
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Sincroniza Productos en Stock"
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
   Begin XtremeSuiteControls.PushButton btnSincroniza 
      Height          =   615
      Index           =   1
      Left            =   4920
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Sincroniza Facturación y Existencias"
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
   Begin XtremeSuiteControls.Label lblStatus 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   6135
      _Version        =   1441793
      _ExtentX        =   10821
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Procesando Sincronización, Este proceso puede durar minutos!"
      BackColor       =   16761024
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
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   735
      Left            =   -120
      TabIndex        =   1
      Top             =   0
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "Ventas En Línea (Mercadito) Sincronización de Catálogos, Inventario y Ventas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "frmInvVentasEnLinea_Sinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnSincroniza_Click(Index As Integer)
Dim pDetalle As String

On Error GoTo vError

lblStatus.Visible = True

Me.MousePointer = vbHourglass

If Index = 0 Then
    pDetalle = "Sincronización de Ventas en Línea: Catálogos de Productos"
    strSQL = "exec spVentas_EnLinea_Sincroniza_Catalogo_Existencias 0"
Else
    pDetalle = "Sincronización de Ventas en Línea: Ventas y Existencias"
    strSQL = "exec spVentas_EnLinea_Sincroniza_Ventas"
End If

Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault

lblStatus.Visible = False

If Not glogon.error Then
    MsgBox "Sincronización realizada satisfactoriamente!", vbInformation
    Call Bitacora("Aplica", pDetalle)
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    lblStatus.Visible = False
    MsgBox "Ocurrió un Error! Verifique que su empresa esté configurada para Ventas en Línea!", vbExclamation

End Sub

Private Sub Form_Load()
On Error GoTo vError
 
 vModulo = 32
 
 Call Formularios(Me)
 
 btnSincroniza(1).Tag = btnSincroniza(0).Tag
 
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub
