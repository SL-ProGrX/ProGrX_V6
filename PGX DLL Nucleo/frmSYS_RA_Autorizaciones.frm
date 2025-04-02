VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSYS_RA_Autorizaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RA Expedientes: Autorizaciones"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   4800
      TabIndex        =   0
      Top             =   2280
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   582
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   3120
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   582
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   570
      Left            =   3120
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   1005
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1095
      Left            =   3120
      TabIndex        =   6
      Top             =   4440
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11451
      _ExtentY        =   1926
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
   Begin XtremeSuiteControls.ComboBox cboHoras 
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   5880
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Autoriza"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Autorizaciones.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtAut_Nombre 
      Height          =   330
      Left            =   4800
      TabIndex        =   11
      Top             =   3720
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   582
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAut_Usuario 
      Height          =   330
      Left            =   3120
      TabIndex        =   12
      ToolTipText     =   "Preisone F4 para Consultar"
      Top             =   3720
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   582
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   330
      Left            =   3120
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   582
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTipo 
      Height          =   330
      Left            =   4800
      TabIndex        =   16
      Top             =   2760
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   582
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   330
      Left            =   5880
      TabIndex        =   17
      Top             =   5880
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de Autorizador"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   18
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario Autorizado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   13
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Horas de Acceso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   2040
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización para Expedientes Restringidos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   1880
      TabIndex        =   5
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Persona Id"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmSYS_RA_Autorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Dim vCodigo As String


Public Sub sbConsulta_Externa(pPersonaId As Long)

Call sbCaso_Load(pPersonaId)

End Sub



Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""

If Len(txtNombre.Text) = 0 Then vMensaje = vMensaje & " - No se Indicó ninguna persona!"
If Len(txtAut_Nombre.Text) = 0 Then vMensaje = vMensaje & " - No se Indicó ningún  Usuario para Autorizar!"
If Len(txtClave.Text) = 0 Then vMensaje = vMensaje & " - No se Indicó la contraseña del autorizador!"


If Len(vMensaje) = 0 Then
    fxVerifica = True
Else
    MsgBox vMensaje, vbExclamation
    fxVerifica = False
End If

End Function



Private Sub btnAutorizacion_Click()
If Not fxVerifica Then
  Exit Sub
End If

On Error GoTo vError

Dim pVence As String

txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)

        
strSQL = "exec spSYS_RA_Autorizacion " & txtCodigo.Text & ", " & cboHoras.Text & ", '" & txtAut_Usuario.Text & "','" _
        & txtNotas.Text & "', '" & glogon.Usuario & "', '" & SIFGlobal.fxStringCifrado(txtClave.Text) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Autorizacion_Id > 0 Then

    Call Bitacora("Registra", "Autorización [" & rs!Autorizacion_Id & "] Expediente Restringido: " & txtCodigo.Text & " Cedula = " & txtCedula)
    
    MsgBox "Registro de Autorización a Expediente Restringido aplicada satisfactoriamente...", vbInformation
    
    UnLoad Me

Else
    MsgBox "No se puede Autorizar, verifique los datos y/o clave de Autorizador!", vbExclamation

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Dim i As Integer

cboHoras.Clear

For i = 1 To 12
cboHoras.AddItem CStr(i)
Next i
cboHoras.Text = "1"

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbCaso_Load(pPersonaId As Long)

On Error GoTo vError


strSQL = "select *, isnull(Fecha_Vence, '2300/01/01') as 'Fecha_Vence_Id' from vSYS_RA_Casos" _
       & " where Persona_Id = " & pPersonaId
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtCodigo.Text = CStr(rs!Persona_Id)
    
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    
    txtTipo.Text = rs!TipoDesc
    txtEstado.Text = rs!EstadoDesc
    txtEstado.Tag = rs!Estado

    
End If
rs.Close

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtAut_Usuario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "USUARIO"
   gBusquedas.Orden = "USUARIO"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT USUARIO, NOMBRE" _
             & " FROM vSYS_RA_Usuarios_Autorizados"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   
   txtAut_Usuario.Text = gBusquedas.Resultado
   txtAut_Nombre.Text = gBusquedas.Resultado2
   
End If
vError:
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAut_Usuario.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "Persona_Id"
   gBusquedas.Orden = "Persona_Id"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT Persona_Id, Cedula, NOMBRE, Estado" _
             & " FROM vSYS_RA_Casos"
   gBusquedas.Filtro = " and cedula like '" & txtCedula & "%'"
   frmBusquedas.Show vbModal
   
   txtCodigo.Text = gBusquedas.Resultado
   If IsNumeric(txtCodigo.Text) Then
      Call sbCaso_Load(txtCodigo.Text)
   End If
   
End If
vError:

End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnAutorizacion.SetFocus
vError:
End Sub

