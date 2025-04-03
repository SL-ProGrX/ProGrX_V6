VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCpr_Solicitud_Autoriza 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autoriza Solicitud"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnResolucion 
      Height          =   732
      Index           =   0
      Left            =   6000
      TabIndex        =   0
      Top             =   1680
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Autorizar"
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
      Picture         =   "frmCpr_Solicitud_Autoriza.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtUEN 
      Height          =   312
      Left            =   4440
      TabIndex        =   2
      ToolTipText     =   "Unidad Estratégica de Negocios"
      Top             =   480
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8700
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnResolucion 
      Height          =   732
      Index           =   1
      Left            =   7560
      TabIndex        =   3
      Top             =   1680
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Rechazar"
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
      Picture         =   "frmCpr_Solicitud_Autoriza.frx":0772
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Resolución para Solicitud de Compra"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   5412
   End
End
Attribute VB_Name = "frmCpr_Solicitud_Autoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'
'Private Sub btnResolucion_Click(Index As Integer)
'
'Dim strSQL As String, vOpcion As String
'
'On Error GoTo vError
'
'If Not fxInvTransaccionesAutoriza(txtCodigo.Text, txtCodigo.Tag, glogon.Usuario) Then
'   MsgBox "Usted no se encuentra Registrado como Autorizados del Usuario que Generó la Transacción...(Verifique)", vbExclamation
'   Exit Sub
'End If
'
'vOpcion = btnResolucion.Item(Index).Caption
'
'If vbNo = MsgBox("Esta seguro que desea " & vOpcion & " la Transaccion en Pantalla", vbYesNo) Then
'   Exit Sub
'End If
'
'
'strSQL = "update pv_InvTranSac set estado = '" & Mid(vOpcion, 1, 1) & "',Autoriza_user = '" & glogon.Usuario _
'       & "',autoriza_fecha = dbo.MyGetdate() where tipo = '" & txtCodigo.Tag _
'       & "' and Boleta = '" & txtCodigo.Text & "'"
'Call ConectionExecute(strSQL)
'
'
'
'Me.MousePointer = vbDefault
'MsgBox "Resolucion Ejecutada Satisfactoriamente...", vbInformation
'
'Unload Me
'
'Exit Sub
'vError:
' Me.MousePointer = vbDefault
' MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'
'
'End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

'txtCodigo.Tag = gInvTran.Tipo
'txtCodigo.Text = gInvTran.Boleta
'txtCausa.Text = gInvTran.Causa

End Sub


