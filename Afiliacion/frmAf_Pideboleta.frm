VERSION 5.00
Begin VB.Form frmAF_Pideboleta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de Reingreso"
   ClientHeight    =   1200
   ClientLeft      =   885
   ClientTop       =   1605
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPromotor 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtboleta 
      Height          =   315
      Left            =   960
      MaxLength       =   5
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblPromotor 
      Caption         =   "Promotor"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblBoleta 
      Caption         =   "Boleta"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   0
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmAF_Pideboleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Devuelve al formulario principal de afiliaciones la boleta y el promotor
'               asignados.
'REFERENCIAS:   ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo CapturaError

   If Trim(cboPromotor) <> "" Then
      frmAF_Principal.cboPromotor = Trim(cboPromotor)
      If Trim(txtBoleta) <> "" Then
         frmAF_Principal.txtBoleta = Trim(txtBoleta)
      End If
      Unload Me
   Else
      MsgBox "Seleccione Un Promotor", vbExclamation, "Faltan Datos"
      cboPromotor.SetFocus
   End If
   
Exit Sub
CapturaError:
    Call ProcedimientoErrores(Me.Name, Err)
   
End Sub


Private Sub cmdCancelar_Click()
'    frmAF_Principal.txtboleta = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    txtBoleta.SetFocus
End Sub


Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verificar y establecer permisos sobre el formulario, y despliega los
'               Promotores activos.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo CapturaError

vModulo = 1
Call Formularios(Me)
Call RefrescaTags(Me)

' Cargamos en la Variable Recordset recPromotor Todos los promotores
' Activos.
strSQL = "Select * from Promotores where Estado=1"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF = False Then
    Do Until rs.EOF
       cboPromotor.AddItem rs!Nombre
       rs.MoveNext
    Loop
    rs.MoveFirst
    cboPromotor.Text = rs!Nombre
End If
rs.Close

Exit Sub
CapturaError:
   Call ProcedimientoErrores(Me.Name, Err)

End Sub

Private Sub txtboleta_KeyPress(KeyAscii As Integer)
On Error GoTo CapturaError

KeyAscii = (Validacion(KeyAscii))

Exit Sub
CapturaError:
   Call ProcedimientoErrores(Me.Name, Err)
End Sub


