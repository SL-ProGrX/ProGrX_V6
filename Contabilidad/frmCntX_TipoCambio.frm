VERSION 5.00
Begin VB.Form frmCntX_TipoCambio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipo de Cambio disponible..."
   ClientHeight    =   3288
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   5748
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3288
   ScaleWidth      =   5748
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMontoFuncional 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   " Monto Moneda Base"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtTC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblTCVariacion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label lblTCPermitido 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblMoneda 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Variación Permitida"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio Permitido"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   5760
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "     Indique el Tipo de Cambio y Monto en Divisa Extranjera que desea registrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   552
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   5892
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto en Divisa Foranea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Cambio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmCntX_TipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

gCntX_TipoCambio.Paso = False

txtTC = gCntX_TipoCambio.TC_Actual
txtMonto = Format(gCntX_TipoCambio.Monto_Actual / fxSys_Tipo_Cambio_Apl(gCntX_TipoCambio.TC_Actual), "Standard")
txtMontoFuncional = Format(gCntX_TipoCambio.Monto_Actual, "Standard")

lblMoneda.Caption = fxCntX_Divisas("D", gCntX_TipoCambio.Moneda)
lblTCPermitido.Caption = "0"
lblTCVariacion.Caption = "0"

strSQL = "select * from CntX_Divisas_Tipo_Cambio where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_divisa = '" & gCntX_TipoCambio.Moneda & "' and '" & Format(gCntX_TipoCambio.fecha, "yyyy/mm/dd") _
       & "' between inicio and corte"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  Select Case fxCntX_Cuenta_Clasificacion(gCntX_TipoCambio.Cuenta)
    Case "A"
       lblTCPermitido.Caption = rs!tc_venta
    Case "P", "C"
       lblTCPermitido.Caption = rs!tc_compra
    Case Else
       lblTCPermitido.Caption = rs!tc_compra
  End Select
  lblTCVariacion.Caption = rs!variacion
 
 If gCntX_TipoCambio.TC_Actual = 0 Then
    txtTC = lblTCPermitido.Caption
 End If
 
End If
rs.Close

vError:

End Sub

Private Sub txtMonto_Change()
On Error GoTo vError

txtMontoFuncional = Format(CCur(txtMonto) * fxSys_Tipo_Cambio_Apl(CCur(txtTC)), "Standard")

vError:
End Sub

Private Function fxVerifica() As Boolean
Dim vMensaje As String

On Error GoTo vError

vMensaje = ""


If Abs(CCur(txtTC.Text) - CCur(lblTCPermitido.Caption)) > CCur(lblTCVariacion.Caption) Then vMensaje = vMensaje & vbCrLf & " - El Tipo de Cambio no es permitido segun variación..."
If CCur(txtTC.Text) = 0 Then vMensaje = vMensaje & vbCrLf & " - El Tipo de Cambio no es válido..."
If CCur(txtMonto.Text) < 0 Then vMensaje = vMensaje & vbCrLf & " - El monto no es válido..."


If Len(vMensaje) > 0 Then
 fxVerifica = False
 MsgBox vMensaje, vbCritical
Else
 fxVerifica = True
End If

Exit Function

vError:
 fxVerifica = False
 MsgBox "Existe Informacion que no se puede calcular...", vbCritical

End Function

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
' Verificar que no exceda la variacion
' Verificar que el monto funcional sea mayor a cero

If KeyCode = vbKeyReturn Then
 If fxVerifica Then
    gCntX_TipoCambio.Monto_Nuevo = fxSys_Tipo_Cambio_Apl(CCur(txtTC)) * CCur(txtMonto)
    gCntX_TipoCambio.TC_Nuevo = CCur(txtTC)
    gCntX_TipoCambio.Paso = True
    Unload Me
 End If
End If

End Sub

Private Sub txtTC_Change()
On Error GoTo vError

txtMontoFuncional = Format(CCur(txtMonto) * fxSys_Tipo_Cambio_Apl(CCur(txtTC)), "Standard")

vError:

End Sub

Private Sub txtTC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 If fxVerifica Then
    gCntX_TipoCambio.Monto_Nuevo = fxSys_Tipo_Cambio_Apl(CCur(txtTC)) * CCur(txtMonto)
    gCntX_TipoCambio.TC_Nuevo = CCur(txtTC)
    gCntX_TipoCambio.Paso = True
    Unload Me
 End If
End If

End Sub
