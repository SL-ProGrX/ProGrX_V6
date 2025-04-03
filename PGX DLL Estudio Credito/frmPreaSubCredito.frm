VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPreaSubCredito 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Solicitud de Crédito"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9540
   Icon            =   "frmPreaSubCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1332
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   9492
      _Version        =   1572864
      _ExtentX        =   16743
      _ExtentY        =   2350
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   612
         Left            =   6840
         TabIndex        =   8
         Top             =   480
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "&Aplicar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   14
         Picture         =   "frmPreaSubCredito.frx":000C
      End
      Begin XtremeSuiteControls.ComboBox cboOperacion 
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Top             =   600
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Operación "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   4212
      _Version        =   1572864
      _ExtentX        =   7435
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCuenta 
      Height          =   312
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   4212
      _Version        =   1572864
      _ExtentX        =   7435
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
      Height          =   312
      Left            =   6840
      TabIndex        =   5
      Top             =   1440
      Width           =   2292
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emitir"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   13
      Left            =   6120
      TabIndex        =   6
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Creación de Solicitud de Crédito"
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
      Height          =   492
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPreaSubCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_Id_Solicitud As String

Public mCedula As String, mLineaCrd As String

Dim vPaso As Boolean, vComite As Integer



Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""

glogon.strSQL = "exec spCRDPreaSolicitudValida '" & gPreAnalisis.Expediente & "'"
If execSql(glogon.strSQL, True) Then
   
    If glogon.Recordset!Aprobado = 0 Then vMensaje = vMensaje & vbCrLf & " - El Expediente no se encuentra aprobado"
    If glogon.Recordset!Pendiente > 0 Then vMensaje = vMensaje & vbCrLf & " - La Solicitud de Crédito ya fue realizada"
    If glogon.Recordset!Maestro = 0 Then vMensaje = vMensaje & vbCrLf & " - Este es un SubExpediente, verifique..."
    If glogon.Recordset!Comite = 0 Then
        vMensaje = vMensaje & vbCrLf & " - No está asignado un comite a evaluación para el expediente"
    Else
        vComite = glogon.Recordset!Comite
    End If
Else
  vMensaje = "Error de comunicación"
End If

If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  fxValida = False
Else
  fxValida = True
End If

End Function


Private Sub cmdAplicar_Click()

On Error GoTo vError

'Validar
If Not fxValida Then Exit Sub

'Aplicar
glogon.strSQL = "exec  spCRDPreaSolicitudCrd '" & gPreAnalisis.Expediente & "'," & vComite _
       & "," & cboBanco.ItemData(cboBanco.ListIndex) & "," & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
       & ",'" & fxTipoDocumento(cboTipoDocumento.Text) & "','" & cboCuenta.ItemData(cboCuenta.ListIndex) _
       & "','','" & glogon.Usuario & "', " & cboOperacion.ItemData(cboOperacion.ListIndex)

If execSql(glogon.strSQL, True) Then
 vModulo = 3
 Call Bitacora("Registra", "Recepción de la Operacion : " & glogon.Recordset!Operacion)
 
 m_Id_Solicitud = glogon.Recordset!Operacion
 
 
  
 Call sbTrazabilidad_Inserta("01", CStr(m_Id_Solicitud), CStr(m_Id_Solicitud))
 
 MsgBox "Solicitud de Credito Generada Satisfactoriamente, Solicitud # " & glogon.Recordset!Operacion
    'Call sbActualizarEstadoPreanalisis
Else
 MsgBox "Se produjo un error"
End If


Exit Sub
 
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
Private Sub sbActualizarEstadoPreanalisis()
Dim StrUpdate1 As String
Dim StrSet As String
Dim vExpediente As String
On Error GoTo vError

StrUpdate1 = "Update CRD_PREA_PREANALISIS SET "

If InStr(1, gPreAnalisis.Expediente, "-", vbTextCompare) > 0 Then
    vExpediente = fxDeCodificaPrimaryKey(gPreAnalisis.Expediente, 1, "-")
Else
    vExpediente = gPreAnalisis.Expediente
End If

StrSet = "ESTADO = " & fxFormatearValor("A", caracter)
StrSet = StrSet & ", USUARIO_GESTION = " & fxFormatearValor(glogon.Usuario, caracter) & ",  FECHA_GESTION = dbo.MyGetdate()"

StrUpdate1 = StrUpdate1 & StrSet & " where COD_PREANALISIS = " & fxFormatearValor(vExpediente, caracter)

StrUpdate1 = StrUpdate1 & " or COD_PREANALISIS_REF = " & fxFormatearValor(vExpediente, caracter)

    
If Not execSql(StrUpdate1, False) Then
    MsgBox "El estado del preanalisis no fue actualizado.", vbInformation, gMsgTitulo
    'Unload Me
End If


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub
Private Sub Form_Load()

Me.Caption = "Expediente : " & gPreAnalisis.Expediente

    m_Id_Solicitud = Empty

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

Call sbCargaCombos

End Sub


Private Sub sbCargaCombos()
Dim strSQL As String, rs As New ADODB.Recordset


vPaso = True


strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

strSQL = "exec spCrdPrea_Operacion_Vincular '" & gPreAnalisis.Expediente & "'"
Call sbCbo_Llena_New(cboOperacion, strSQL, False, True)


cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("RE")
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("ND")

cboTipoDocumento.Text = fxTipoDocumento("TE")


vPaso = False

Call cboBanco_Click

End Sub


Function fxTipoDocumento(vTipo As String) As String
Select Case vTipo
 Case "CK"
   fxTipoDocumento = "Cheque"
 Case "TE"
   fxTipoDocumento = "Transferencia"
 Case "EF", "RE"
   fxTipoDocumento = "Efectivo"
 Case "ND"
   fxTipoDocumento = "Nota Debito"
 Case "NC"
   fxTipoDocumento = "Nota Credito"
 Case "OT"
   fxTipoDocumento = "Otro..."
'-------
 Case "Cheque"
   fxTipoDocumento = "CK"
 Case "Transferencia"
   fxTipoDocumento = "TE"
 Case "Efectivo"
   fxTipoDocumento = "EF"
 Case "Nota Debito"
   fxTipoDocumento = "ND"
 Case "Nota Credito"
   fxTipoDocumento = "NC"
 Case "Otro..."
   fxTipoDocumento = "OT"
 Case Else
   fxTipoDocumento = ""
End Select
End Function

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & gPreAnalisis.Tag1 & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

