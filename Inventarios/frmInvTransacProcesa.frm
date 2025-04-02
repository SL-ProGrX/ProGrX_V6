VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmInvTransacProcesa 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Procesa Transacciones"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   732
      Left            =   7800
      TabIndex        =   0
      Top             =   1680
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Procesar"
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
      Picture         =   "frmInvTransacProcesa.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2640
      TabIndex        =   2
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
   Begin XtremeSuiteControls.FlatEdit txtCausa 
      Height          =   312
      Left            =   4440
      TabIndex        =   3
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
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesa en Inventarios la Boleta y realiza todas las afectaciones al mismo."
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
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmInvTransacProcesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()
Dim strSQL As String

On Error GoTo vError

'Verificar la Fecha de Afectacion si es válida o no
If Not fxInvPeriodos(gInvTran.fecha) Then
 MsgBox " - El periodo de la Transaccion ya fue cerrado o no es válido ..."
 Exit Sub
End If

If vbNo = MsgBox("Esta seguro que desea Procesar la Transaccion en Pantalla", vbYesNo) Then
   Exit Sub
End If


Me.MousePointer = vbHourglass

strSQL = "exec spINVTranProcesa '" & txtCodigo.Tag & "','" & txtCodigo.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Procesamiento Finalizado Satisfactoriamente...", vbInformation
Unload Me

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

txtCodigo.Tag = gInvTran.Tipo
txtCodigo.Text = gInvTran.Boleta
txtCausa.Text = gInvTran.Causa

End Sub

