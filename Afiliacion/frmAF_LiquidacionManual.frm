VERSION 5.00
Begin VB.Form frmAF_LiquidacionManual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación Manual"
   ClientHeight    =   2460
   ClientLeft      =   1560
   ClientTop       =   3735
   ClientWidth     =   6015
   ControlBox      =   0   'False
   HelpContextID   =   1006
   Icon            =   "frmAF_LiquidacionManual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDisponible 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtAtrasado 
      Height          =   315
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtAmortizacion 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtDeuda 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtIntC 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtAbono 
      Height          =   315
      Left            =   4080
      MaxLength       =   15
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtIntM 
      Height          =   315
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtSaldo 
      Height          =   315
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtMonto 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtSolicitud 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Disponible"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Total Atrasado"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Amortización"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Total Deuda"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Código"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   18
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6240
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label6 
      Caption         =   "Abono"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Int. Mora."
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Int. Cor."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Saldo"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Monto Apr."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "# Operación"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   0
      X2              =   6240
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmAF_LiquidacionManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Corrobora que el abono no exceda al monto por asignar, y lo asigna a la
'               pantalla anterior, afectando en pantalla el saldo resultante del credito.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim curIntc As Currency
Dim curIntm As Currency
Dim curAbono As Currency
Dim curSaldo As Currency

Me.MousePointer = vbHourglass

If Trim(txtAbono) = "" Then
  MsgBox "Registre El Monto Del Abono", vbExclamation, "Faltan Datos"
Else
  curAbono = CCur(txtAbono)
  
  If curAbono > GLOBALES.gcurMontoLibre Then
     MsgBox "El Monto Del Abono es Superior" & vbCrLf & _
     "Al Monto Del Ahorro Disponible" & vbCrLf & _
     "Dispone Unicamente de " & Format(GLOBALES.gcurMontoLibre, "Standard"), _
     vbExclamation, "Verifique El Abono"
     Me.MousePointer = vbDefault
     Exit Sub
  Else
     GLOBALES.gcurMontoAbonado = curAbono
  End If
  
  frmAF_Liquidacion.F1Creditos.Col = 13 '13
  frmAF_Liquidacion.F1Creditos.Text = Format(curAbono, "Standard")
  
  frmAF_Liquidacion.F1Creditos.Col = 9
  If Trim(frmAF_Liquidacion.F1Creditos.Text) <> "" Then
     curIntc = CCur(frmAF_Liquidacion.F1Creditos.Text)
  Else
     curIntc = 0
  End If
  
  frmAF_Liquidacion.F1Creditos.Col = 10
  If Trim(frmAF_Liquidacion.F1Creditos.Text) <> "" Then
     curIntm = CCur(frmAF_Liquidacion.F1Creditos.Text)
  Else
     curIntm = 0
  End If
    
  curAbono = curAbono - curIntc
  curAbono = curAbono - curIntm
  
  If curAbono > 0 Then
     frmAF_Liquidacion.F1Creditos.Col = 8
     curSaldo = CCur(frmAF_Liquidacion.F1Creditos.Text)
     
     frmAF_Liquidacion.F1Creditos.Col = 12
     frmAF_Liquidacion.F1Creditos.Text = Format((curSaldo - curAbono), "Standard")
  Else
     frmAF_Liquidacion.F1Creditos.Col = 8
     curSaldo = CCur(frmAF_Liquidacion.F1Creditos.Text)
     
     frmAF_Liquidacion.F1Creditos.Col = 12
     frmAF_Liquidacion.F1Creditos.Text = Format(curSaldo, "Standard")
  End If
  
  Unload Me
End If

Me.MousePointer = vbDefault

End Sub

Private Sub cmdCancelar_Click()
GLOBALES.gcurMontoAbonado = 0
Unload Me
End Sub

Private Sub Form_Activate()
txtDeuda = Format((CCur(txtSaldo) + (CCur(txtIntC) + CCur(txtIntM))), "Standard")
txtDisponible = Format(GLOBALES.gcurMontoLibre, "Standard")
End Sub

Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verificar y establecer permisos sobre el formulario.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtAbono_Change()
Set GLOBALES.gCajaTxt = txtAbono
Call ValidaMonto
End Sub

Private Sub txtAbono_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   cmdAceptar.SetFocus
End If
End Sub


Private Sub txtAbono_LostFocus()
Dim curAbono As Currency
Dim curIntc As Currency
Dim curIntm As Currency
Dim curSaldo As Currency

If Trim(txtSaldo) <> "" Then
   curSaldo = CCur(txtSaldo)
Else
   curSaldo = 0
End If

If Trim(txtIntC) <> "" Then
   curIntc = CCur(txtIntC)
Else
   curIntc = 0
End If

If Trim(txtIntM) <> "" Then
   curIntm = CCur(txtIntM)
Else
   curIntm = 0
End If

If Trim(txtAbono) <> "" Then
   curAbono = CCur(txtAbono)
Else
   curAbono = 0
End If

curSaldo = curSaldo + curIntc + curIntm

If curAbono > curSaldo Then
   MsgBox "Exceso En El Monto Del Abono", vbExclamation, "Verifique El Abono"
   txtAbono = ""
   txtAbono.SetFocus
End If

End Sub


Private Sub txtAmortizacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtAtrasado.SetFocus
End If
End Sub


Private Sub txtAtrasado_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtDeuda.SetFocus
End If
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtMonto.SetFocus
End If
End Sub


Private Sub txtDeuda_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtAbono.SetFocus
End If
End Sub


Private Sub txtIntC_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtIntM.SetFocus
End If
End Sub


Private Sub txtIntM_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtAmortizacion.SetFocus
End If
End Sub


Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtSaldo.SetFocus
End If
End Sub


Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtIntC.SetFocus
End If
End Sub


Private Sub txtSolicitud_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtCodigo.SetFocus
End If
End Sub


