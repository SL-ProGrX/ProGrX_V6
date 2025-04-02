VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmTES_Reposicion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reposición"
   ClientHeight    =   7164
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7164
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   7200
      Top             =   240
   End
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox txtVerifica 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   4440
      Width           =   3735
   End
   Begin VB.TextBox txtContraseña 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   25
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   24
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox txtCuenta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtTipoCaso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtID_Banco 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtBanco 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   5415
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtTipo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtBeneficiario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox txtNumeroSolicitud 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtDetalle 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3240
      Width           =   6615
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin XtremeSuiteControls.PushButton cmdReponer 
      Height          =   792
      Left            =   5880
      TabIndex        =   35
      Top             =   6240
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   1397
      _StockProps     =   79
      Caption         =   "&Marca para Reposición"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   16
      Picture         =   "frmTES_Reposicion.frx":0000
   End
   Begin VB.Label lblDocumento 
      Alignment       =   1  'Right Justify
      Caption         =   "No. Documento:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   34
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "No. Documento:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   33
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Notas de Reposición"
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
      Height          =   252
      Index           =   4
      Left            =   4080
      TabIndex        =   32
      Top             =   4200
      Width           =   3612
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Verificación "
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
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   30
      Top             =   4200
      Width           =   3492
   End
   Begin VB.Label Label3 
      Caption         =   "Clave"
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
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Autorización"
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
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario"
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
      Left            =   240
      TabIndex        =   27
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "No.Cuenta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   10
      Left            =   4440
      TabIndex        =   22
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Caso"
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
      Index           =   9
      Left            =   360
      TabIndex        =   20
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Detalle"
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
      Left            =   360
      TabIndex        =   19
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Banco"
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
      Left            =   360
      TabIndex        =   18
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Emisión"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4440
      TabIndex        =   17
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Doc."
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
      Index           =   3
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Monto"
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
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Beneficiario"
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
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Código"
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
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Index           =   7
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "No. Solicitud"
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
      Index           =   6
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REPOSICION DE PAGO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4332
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7800
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmTES_Reposicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReponer_Click()
Dim strSQL As String, rs As New ADODB.Recordset


If txtVerifica.Tag <> "S" Then
   MsgBox "Identifique las notas de la verificación antes de proceder...!!!", vbExclamation
   Exit Sub
End If

'Verificar Usuarios y Claves de Autorización
strSQL = "select isnull(count(*),0) as Existe from tes_autorizaciones where nombre = '" _
       & txtUsuario & "' and estado = 'A' and clave = '" & fxTESCifrado(txtContraseña) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  rs.Close
  MsgBox "El usuario y clave de autorización no concuerda con ninguno de los registrados, verifique...", vbExclamation
  Exit Sub
End If
rs.Close

strSQL = MsgBox("Confirma Reposición?", vbExclamation + vbYesNo + vbDefaultButton2)
If strSQL = vbNo Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "Exec spTES_Reposicion " & txtNumeroSolicitud.Text & ",'" & glogon.Usuario & "', '" & txtUsuario.Text & "','" & Mid(txtNotas.Text, 1, 250) & "'"
Call ConectionExecute(strSQL)

Call sbTesBitacoraEspecial(txtNumeroSolicitud.Text, "18", Mid(txtNotas.Text, 1, 150))
Call Bitacora("Aplica", "Reposición de Pago - Solicitud :" & txtNumeroSolicitud.Text)

MsgBox "Reposición Registrada Satifactoriamente!", vbInformation

txtUsuario = ""
txtContraseña = ""

Me.MousePointer = vbDefault

Unload Me

Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 9

End Sub

Private Sub Form_Load()

vModulo = 9

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

TimerX.Interval = 0

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select C.Nsolicitud,C.Codigo,C.beneficiario,C.tipo,C.estado,C.ndocumento,C.id_banco,B.descripcion as BancoX" _
       & ",T.descripcion as TipoDocX,C.Monto,C.Fecha_Emision,C.Tipo_Beneficiario, C.Cta_Ahorros" _
       & ",C.Detalle1 + ' ' + C.Detalle2 + ' ' + isnull(C.Detalle3 ,'') + ' ' + isnull(C.Detalle4 ,'')  + ' ' + isnull(C.Detalle5 ,'') as 'Detalle'" _
       & ", case when C.Tipo_Beneficiario = 1 then 'Personas'" _
       & " when C.Tipo_Beneficiario = 2 then 'Bancos'" _
       & " when C.Tipo_Beneficiario = 3 then 'Proveedores'" _
       & " when C.Tipo_Beneficiario = 4 then 'Acreedores' end as 'TipoBeneficiario'" _
       & ", isnull(C.REPOSICION_IND,0) as 'ReposicionPaso' " _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_Banco" _
       & " inner join tes_tipos_doc T on C.tipo = T.tipo" _
       & " inner join tes_banco_docs Y on C.id_banco = Y.id_Banco and C.tipo = Y.tipo" _
       & " where C.nsolicitud = " & GLOBALES.gTag & " and C.estado in('T','E','I')"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
    'Seccion de Verificación
    txtVerifica.Tag = "N"
    txtVerifica.Text = "Este documento no es valido para reposición..."
    rs.Close
    Exit Sub

Else
    If rs!ReposicionPaso = 1 Then
        'Seccion de Verificación
        txtVerifica.Tag = "N"
        txtVerifica.Text = "Este documento ya Registró Reposición Anteriormente!..."
        rs.Close
        Exit Sub
    End If

    txtNumeroSolicitud.Text = rs!NSolicitud
    txtEstado.Text = "Emitido"
    txtCodigo.Text = rs!Codigo
    txtBeneficiario.Text = rs!Beneficiario
    
    txtFecha.Text = Format(rs!Fecha_Emision & "", "dd/mm/yyyy")
    lblDocumento.Caption = rs!nDocumento & ""
    
    txtBanco.Text = rs!BancoX
    txtID_Banco.Text = rs!id_Banco
    
    txtMonto.Text = Format(rs!Monto, "Standard")
    txtTipo.Text = rs!TipoDocX
    txtTipoCaso.Text = rs!TipoBeneficiario
    txtTipoCaso.Tag = rs!Tipo_Beneficiario
    
    txtDetalle.Text = rs!Detalle & ""

End If
rs.Close

'Seccion de Verificación
txtVerifica.Tag = "S"


If txtTipoCaso.Tag <> "3" Then
 txtVerifica = txtVerifica & vbCrLf & " - El Tipo de Beneficiario no aplica (Solo Pago de Proveedores)..."
 txtVerifica.Tag = "N"
End If



'Fin de Verificacion


If txtVerifica.Tag = "S" Then
   txtVerifica.Text = "----> Este Documento se puede marcar para reponer"
   txtVerifica.ForeColor = vbBlue
Else
   txtVerifica.ForeColor = vbRed
End If

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 txtVerifica.Text = fxSys_Error_Handler(Err.Description)
 txtVerifica.ForeColor = vbRed
 txtVerifica.Tag = "N"
 
End Sub
