VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmPosCajaApertura 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura de Cajas"
   ClientHeight    =   6000
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7788
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   7788
   Begin XtremeSuiteControls.GroupBox gbApertura 
      Height          =   3252
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   7812
      _Version        =   1245187
      _ExtentX        =   13779
      _ExtentY        =   5736
      _StockProps     =   79
      Caption         =   "Apertura de Caja: "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin VB.TextBox txtEA_EfectivoInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtEA_Cierre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtEA_CierreNumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5160
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtEA_CxCInternaInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtEA_CxCInternaApertura 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtEA_CxCExternaInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtEA_CxCExternaApertura 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
      Begin XtremeSuiteControls.PushButton cmdApertura 
         Height          =   492
         Left            =   3120
         TabIndex        =   22
         Top             =   2520
         Width           =   4092
         _Version        =   1245187
         _ExtentX        =   7218
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Realizar Apertura de Caja"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmPosCajaApertura.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   7812
         _Version        =   1245187
         _ExtentX        =   13779
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Apertura de Caja"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicial en Efectivo"
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
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   2772
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo Cierre Registrado"
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
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   2772
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ultimo Estado"
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
         Height          =   312
         Index           =   5
         Left            =   3120
         TabIndex        =   19
         Top             =   600
         Width           =   2052
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado de Apertura"
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
         Height          =   312
         Index           =   6
         Left            =   5160
         TabIndex        =   18
         Top             =   600
         Width           =   2052
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Doc x Cobrar Internos"
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
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   2892
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Doc x Cobrar Externos"
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
         Height          =   252
         Index           =   8
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   3132
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2760
      TabIndex        =   1
      Top             =   1560
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
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
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Realice la Apertura de su Caja para Iniciar su actividad y movimientos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPosCajaApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxExisteApertura(vCaja As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from pv_cajas_ac where usuario = '" _
       & glogon.Usuario & "' and cod_caja = '" & vCaja & "' and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   fxExisteApertura = True
Else
   fxExisteApertura = False
End If
rs.Close

End Function



Private Sub cbo_Click()
Call sbLimpiaDatos
gbApertura.Enabled = False

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or vbKeyTab Then txtClave.SetFocus
End Sub

Private Sub cmdApertura_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long

On Error GoTo vError

If fxExisteApertura(cbo.ItemData(cbo.ListIndex)) Then
  MsgBox "Ya existe una apertura para esta caja...", vbExclamation
  txtUsuario.SetFocus
  gbApertura.Enabled = False
  Exit Sub
End If

If CCur(txtEA_EfectivoInicial) > CCur(txtEfectivo) Then
  MsgBox "El Monto de Efectivo Inicial, no es válido ya que no es mayor que el actual...", vbExclamation
  txtEfectivo.SetFocus
  Exit Sub
End If

Me.MousePointer = vbHourglass
 
'Saca el Consecutivo de Gestion Apertura/Cierre
strSQL = "select isnull(max(cod_ac),0) as Ultimo from pv_cajas_ac where usuario = '" _
       & glogon.Usuario & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
  i = rs!ultimo + 1
rs.Close

strSQL = "insert pv_cajas_ac(cod_ac,cod_caja,usuario,estado,ap_fecha,ap_saldo_efectivo" _
       & ",ap_saldo_docint,ap_saldo_docext,ci_saldo_efectivo,ci_saldo_docint,ci_saldo_docext)" _
       & " values(" & i & ",'" & cbo.ItemData(cbo.ListIndex) & "','" & glogon.Usuario & "','A',dbo.MyGetdate()," _
       & CCur(txtEfectivo) & "," & CCur(txtEA_CxCInternaApertura) & "," & CCur(txtEA_CxCExternaApertura) & "," _
       & CCur(txtEfectivo) & "," & CCur(txtEA_CxCInternaApertura) & "," & CCur(txtEA_CxCExternaApertura) & ")"
Call ConectionExecute(strSQL)

strSQL = "update pv_cajas set ult_apertura = dbo.MyGetdate(),saldo_efectivo = " & CCur(txtEfectivo) _
       & ",saldo_documentos = " & CCur(txtEA_CxCExternaApertura) + CCur(txtEA_CxCInternaApertura) _
       & " where usuario = '" & glogon.Usuario & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Apertura Num." & i & " Caja:" & cbo.ItemData(cbo.ListIndex) & ".US:" & glogon.Usuario)

Me.MousePointer = vbDefault

MsgBox "Apertura # " & i & " registrada satisfactoriamente...", vbInformation
Call sbLimpiaDatos

txtUsuario.SetFocus
gbApertura.Enabled = False

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 33

'Solo carga las cajas que no tengan aperturas abiertas y que no esten bloqueadas
'del usuario activo
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtUsuario = glogon.Usuario
txtClave = ""
strSQL = "select rtrim(Cd.cod_caja) as 'IdX' , rtrim(Cd.nombre) as 'ItmX'" _
       & " from pv_cajas Cd" _
       & " where Cd.estado = 'A' and Cd.usuario = '" & glogon.Usuario & "' and Cd.Bloqueo = 0" _
       & " and dbo.fxPOS_Caja_Apertura_Existe(Cd.cod_Caja, Cd.Usuario) = 0"

Call sbCbo_Llena_New(cbo, strSQL, False, True)

End Sub

Private Sub sbLimpiaDatos()
    txtEA_Cierre = ""
    txtEA_CierreNumero = 0
    txtEA_CxCExternaApertura = "0.00"
    txtEA_CxCExternaInicial = "0.00"
    txtEA_CxCInternaApertura = "0.00"
    txtEA_CxCInternaInicial = "0.00"
    txtEA_EfectivoInicial = "0.00"
    txtEfectivo = "0.00"
End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

'Verifica datos nuevamente por razones de seguridad por violación
'por concurrencias.

'1. Verificar que el Estado este Activa / en el Load esta validado
'2. Que no se encuentre Bloqueada
'3. Verificar si la caja esta abierta (Apertura) y Sacar el Consecutivo
'   de la apertura.

If KeyCode = vbKeyReturn Then
 strSQL = "select bloqueo from pv_cajas where usuario = '" _
        & txtUsuario & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "' and clave = '" _
        & fxPosEncrypta(txtClave) & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs.EOF And rs.BOF Then
   MsgBox "Caja: verifique su Usuario y Clave para Esta Caja ...", vbExclamation
 Else
  If rs!bloqueo = 0 Then
     gCajas.Caja = cbo.ItemData(cbo.ListIndex)
     gCajas.Usuario = txtUsuario
     
     'Revisa que la caja no sea nueva, de lo contrario por defecto puede registrar la
     'Apertura...
     rs.Close
     strSQL = "select isnull(max(cod_ac),0) as UltCierre from pv_cajas_ac where cod_caja = '" & gCajas.Caja _
            & "' and usuario = '" & gCajas.Usuario & "' and estado = 'C'"
     Call OpenRecordSet(rs, strSQL)
     
     gCajas.Apertura = rs!ultCierre
     gbApertura.Enabled = True
     
     If rs!ultCierre = 0 Then
        Call sbLimpiaDatos
        txtEfectivo.SetFocus
     Else
        strSQL = "select * from pv_cajas_ac where cod_caja = '" & gCajas.Caja _
               & "' and usuario = '" & gCajas.Usuario & "' and cod_ac = " & rs!ultCierre
        rs.Close
        Call OpenRecordSet(rs, strSQL)
        txtEA_Cierre = rs!CI_Fecha & ""
        txtEA_CierreNumero = rs!cod_ac
        txtEA_CxCExternaApertura = Format(rs!ci_saldo_docext, "Standard")
        txtEA_CxCExternaInicial = Format(rs!ci_saldo_docext, "Standard")
        txtEA_CxCInternaApertura = Format(rs!ci_saldo_docint, "Standard")
        txtEA_CxCInternaInicial = Format(rs!ci_saldo_docint, "Standard")
        txtEA_EfectivoInicial = Format(rs!ci_saldo_efectivo, "Standard")
        txtEfectivo = txtEA_EfectivoInicial
        txtEfectivo.SetFocus
     End If
  
  Else
    MsgBox "La Caja se encuentra Bloqueada...", vbExclamation
  
  End If 'Bloqueo
 
 End If 'Select cajas
 rs.Close

End If

End Sub

Private Sub txtEfectivo_GotFocus()
On Error GoTo vError
  txtEfectivo = CCur(txtEfectivo)
  txtEfectivo.SelStart = Len(txtEfectivo)
vError:
End Sub

Private Sub txtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdApertura.SetFocus
End Sub

Private Sub txtEfectivo_LostFocus()
On Error GoTo vError
  txtEfectivo = Format(CCur(txtEfectivo), "Standard")
vError:
End Sub
