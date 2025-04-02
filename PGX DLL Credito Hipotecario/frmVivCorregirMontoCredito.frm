VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmVivCorregirMontoCredito 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Corregir monto del crédito"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3836
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
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTasa 
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPlazo 
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMontoNoGravable 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3836
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
      Text            =   "0.00"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdAceptar 
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Aceptar"
      BackColor       =   16777215
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
      Picture         =   "frmVivCorregirMontoCredito.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdCancelar 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cancelar"
      BackColor       =   16777215
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
      Picture         =   "frmVivCorregirMontoCredito.frx":07D8
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3836
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
      Text            =   "0.00"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuota"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tasa"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Plazo"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   975
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   3255
      _Version        =   1310723
      _ExtentX        =   5741
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "El monto no gravable, se rebajará del monto del crédito para efectos de cálculos de honorarios y avalúos "
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Monto No Gravable"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Monto del Crédito"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVivCorregirMontoCredito.frx":0FA5
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5145
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmVivCorregirMontoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Public m_Operacion As Long
Private m_cambioDatos As Boolean




Private Sub sbCargarOperacion()

On Error GoTo vError

strSQL = "select dbo.MyGetdate() as Fecha, S.nombre,C.descripcion as CodDesc,X.descripcion as ComDesc," _
       & "R.cedula,R.codigo,R.id_solicitud,R.id_comite,R.Amortiza,R.saldo,R.estadosol,R.estado," _
       & "R.montoapr,R.ts,R.cuota,R.plazo,R.int,R.interesv,R.interesc,R.montosol," _
       & "R.observacion,R.pagare,R.fechaforp,R.fechasol,R.fechares,R.ind_deduce_planilla," _
       & "R.proceso,R.garantia,R.documento_referido,R.primer_cuota,R.emitir,R.prideduc," _
       & "R.userrec,R.userres,R.userfor,R.usertesoreria,R.cod_banco,R.CTA_BANCO,R.Garantia_Fnd" _
       & ",R.fecha_inicio_calculo,R.cod_grupo,R.cod_destino,R.id_comite,R.Autoriza_User,R.Autoriza_Fecha" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Comites X on R.id_comite = X.id_comite" _
       & " where R.id_solicitud = " & gOperacion
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
     txtCuota.Text = Format(IIf(IsNull(rs!cuota), 0, rs!cuota), "Standard")
     txtPlazo.Text = IIf(IsNull(rs!Plazo), 0, rs!Plazo)
     txtTasa.Text = Format(IIf(IsNull(rs!Int), 0, rs!Int), "Standard")
     txtMonto.Text = Format(IIf(IsNull(rs!montosol), 0, rs!montosol), "Standard")
     rs.Close
End If

m_cambioDatos = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbMaxMontoNoGravable()

On Error GoTo vError

strSQL = "select isnull(max(MontoNoGravable),0) as MontoNoGravable" _
       & " from viviendaGarantia" _
       & " where NumeroOperacion = " & gOperacion

Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    txtMontoNoGravable.Text = Format(rs!MontoNoGravable, "Standard")
    rs.Close
Else
    txtMontoNoGravable.Text = Format(0, "Standard")
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdAceptar_Click()
Dim vEstado As String

vEstado = ""
If m_cambioDatos = False Then Exit Sub

vEstado = ObjConsultar.fxEstadoOperacion(gOperacion)

If Not ((vEstado = "R") Or (vEstado = "P")) Then
    Me.MousePointer = vbDefault
    MsgBox "Para modificar el monto del crédito la operación debe estar en estado Recibida o Pendiente.", vbExclamation
    Exit Sub
End If

If Val(txtMonto.Text) = 0 Then
    Me.MousePointer = vbDefault
    MsgBox "El monto de la operación es inválido.", vbExclamation
    txtMonto.SetFocus
    Exit Sub
End If


If (MsgBox("¿ Confirma que desea modificar el monto del crédito.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub
strSQL = "Update Reg_creditos set montoapr = " & CCur(txtMonto.Text) & ", montosol = " & CCur(txtMonto.Text) _
        & ",saldo = " & CCur(txtMonto.Text) & ", cuota = " & CCur(txtCuota.Text) _
        & " where  id_solicitud = " & gOperacion
Call ConectionExecute(strSQL)
If Not glogon.error Then
    MsgBox "Información fue actualizada correctamente.", vbInformation
End If

strSQL = "Update ViviendaGarantia set MontoNoGravable = " & CCur(txtMontoNoGravable.Text) _
       & " where NumeroOperacion = " & gOperacion
Call ConectionExecute(strSQL)
If Not glogon.error Then
    Call Bitacora("APLICA", "Actualiza monto no gravable operación: " & gOperacion)
End If

End Sub

Private Sub cmdCancelar_Click()
m_cambioDatos = False
UnLoad Me
End Sub

Private Sub Form_Load()
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

'' Carga nombre de la ternimal
If Len(glogon.Maquina) = 0 Then
    Call sbMaquina
End If

Call sbCargarOperacion
Call sbMaxMontoNoGravable

m_cambioDatos = False

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo vError

Select Case Me.ActiveControl.Name

Case "txtMonto"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtMonto.Text), KeyAscii)
Case "txtMontoNoGravable"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtMontoNoGravable.Text), KeyAscii)
    
End Select


Exit Sub

vError:
MsgBox "Ocurrió un error validar la información de los formatos. " & "-" & Err.Description, vbExclamation

End Sub


Private Sub txtMonto_Change()
On Error GoTo vError
m_cambioDatos = True

If Not IsNumeric(txtMonto.Text) Then Exit Sub

If CCur(IIf((txtTasa.Text = ""), 0, txtTasa.Text)) > 0 And CCur(IIf((txtPlazo.Text = ""), 0, txtPlazo.Text)) > 0 _
    And CCur(IIf((txtMonto.Text = ""), 0, txtMonto.Text)) > 0 Then
 txtCuota.Text = fxCalcula_Cuota(CCur(txtMonto.Text), CCur(txtPlazo.Text), CCur(txtTasa.Text))
End If

vError:

End Sub
Private Sub txtMonto_GotFocus()
If Not IsNumeric(txtMonto.Text) Then Exit Sub
txtMonto.Text = CCur(txtMonto.Text)

End Sub

Private Sub txtMonto_LostFocus()
If Not IsNumeric(txtMonto.Text) Then Exit Sub
txtMonto.Text = Format(txtMonto.Text, "Standard")

End Sub

Private Sub txtMontoNoGravable_Change()
m_cambioDatos = True
End Sub

Private Sub txtMontoNoGravable_GotFocus()
If Not IsNumeric(txtMontoNoGravable.Text) Then Exit Sub
    m_cambioDatos = True
    txtMontoNoGravable.Text = CCur(txtMontoNoGravable.Text)
End Sub

Private Sub txtMontoNoGravable_LostFocus()

If Not IsNumeric(txtMontoNoGravable.Text) Then
txtMontoNoGravable.Text = Format(0, "Standard")
End If

txtMontoNoGravable.Text = Format(txtMontoNoGravable.Text, "Standard")
End Sub
