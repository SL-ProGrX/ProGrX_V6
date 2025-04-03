VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPrea_Edad_Justificacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Justificación"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnAceptar 
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Aceptar"
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
      Appearance      =   21
      Picture         =   "frmPrea_Edad_Justificacion.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtJustificación 
      Height          =   3015
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   7695
      _Version        =   1572864
      _ExtentX        =   13573
      _ExtentY        =   5318
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.UpDown UpDownCuotas 
      Height          =   315
      Left            =   2865
      TabIndex        =   4
      Top             =   4080
      Width           =   270
      _Version        =   1572864
      _ExtentX        =   476
      _ExtentY        =   556
      _StockProps     =   64
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Value           =   100
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtS_Privado_Porc"
      BuddyProperty   =   ""
   End
   Begin XtremeSuiteControls.FlatEdit txtCuotas 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   4080
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "100"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de Cuotas"
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
      Height          =   330
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _Version        =   1572864
      _ExtentX        =   14420
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Justificación por Incumplimiento de Edad de Pensión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "frmPrea_Edad_Justificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mExpediente As String

Private Sub sbConsulta()

On Error GoTo vError

strSQL = "SELECT ISNULL(APL_JUSTIFICACION_EDAD,0) 'EDAD_APLICA'" _
       & ", ISNULL(JUSTIFICACION_EDAD,'') 'EDAD_JUSTIFICACION'" _
       & ", ISNULL(CANTIDAD_CUOTAS_JUSTI_EDAD, PLAZO) AS 'EDAD_CUOTAS'" _
       & " FROM CRD_PREA_PREANALISIS WHERE COD_PREANALISIS = '" & mExpediente & "'"
Call OpenRecordSet(rs, strSQL)
 
 txtJustificación.Tag = rs!Edad_Aplica
 txtJustificación.Text = rs!Edad_Justificacion
 txtCuotas.Text = rs!Edad_Cuotas
 UpDownCuotas.Value = rs!Edad_Cuotas
 
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGuardar()

On Error GoTo vError

If Len(txtJustificación.Text) < 50 Then
    MsgBox "Debe ingresar una justificación de al menos 50 caracteres para continuar!", vbExclamation
    Exit Sub
End If


strSQL = "exec spCrdPreaGuardaJustificacionEdadPension '" & mExpediente & "', 1, '" _
       & txtJustificación.Text & "', " & txtCuotas.Text & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If glogon.error Then
    Exit Sub
Else
    UnLoad Me
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAceptar_Click()

Call sbGuardar

End Sub

Private Sub Form_Load()
vModulo = 3

mExpediente = GLOBALES.gTag

txtJustificación.Text = ""
txtJustificación.Tag = 0

txtCuotas.Text = 1
UpDownCuotas.Value = 1

Call sbConsulta

End Sub


Private Sub txtCuotas_Change()
On Error GoTo vError

If Not IsNumeric(txtCuotas.Text) Then
    txtCuotas.Text = "1"
End If

If CLng(txtCuotas.Text) < 0 Then
    txtCuotas.Text = "1"
End If


UpDownCuotas.Value = txtCuotas.Text

Exit Sub

vError:

    txtCuotas.Text = "1"
    UpDownCuotas.Value = txtCuotas.Text

End Sub

Private Sub UpDownCuotas_DownClick()
txtCuotas.Text = UpDownCuotas.Value
End Sub

Private Sub UpDownCuotas_UpClick()
txtCuotas.Text = UpDownCuotas.Value
End Sub
