VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajas_Cierre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierre de Caja"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaldosFavor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtDiferencia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtDocumentosFin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtDocumentosIni 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtEfectivoFin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtEfectivoIni 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtMontoDoc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtDepNum 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtMontoEfectivo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin MSComctlLib.Toolbar tblAplicar 
      Height          =   360
      Left            =   3960
      TabIndex        =   21
      Top             =   3000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      ButtonWidth     =   1931
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesar"
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Aplica la Transacción"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaja_Cierre.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaja_Cierre.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   6600
      X2              =   0
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Cierre de Caja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1200
      TabIndex        =   29
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Saldos a Favor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Caja # :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   26
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Apertura #:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   25
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Efectivo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   24
      Top             =   2400
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Diferencia+/-:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3360
      TabIndex        =   22
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Resultado de Cierre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Documentos Recibido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3240
      TabIndex        =   15
      Top             =   4680
      Width           =   1605
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Efectivo Recibido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Documentos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Documentos Inicial:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3240
      TabIndex        =   11
      Top             =   4200
      Width           =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Efectivo inicial:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   1065
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   6600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Numero Dep:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Detalle de Cierre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label lblFechaApertura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblUsuarioApertura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   2
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label lblCaja 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblApertura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      DrawMode        =   4  'Mask Not Pen
      X1              =   0
      X2              =   6360
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmCajas_Cierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cMontoEfectivo As Currency
Dim cMontoDocumentos As Currency

Private Sub Form_Load()
lblCaja.Caption = ModuloCajas.mCaja
lblApertura.Caption = ModuloCajas.mApertura
lblUsuarioApertura.Caption = glogon.Usuario
Call sbCajaDatos

If ModuloCajas.mTipoCierre = "C" Then
    Me.Height = 3810
Else
    Me.Height = 5985
End If

End Sub


Private Sub sbCajaDatos()
Dim strSQL As String, rs As New ADODB.Recordset

'strSQL = " select APERTURA_FECHA,SI_EFECTIVO,SI_DOCUMENTOS from CAJAS_APERTURAS_CIERRES where cod_caja = '" & ModuloCajas.mCaja & "'  and cod_apertura = " & ModuloCajas.mApertura & ""
'rs.Open strSQL, glogon.Conection, adOpenStatic
'If Not rs.EOF Then
'  lblFechaApertura.Caption = Format(rs!APERTURA_FECHA, "dd/mm/yyyy")
'  txtEfectivoIni.Text = Format(rs!SI_EFECTIVO, "Standard")
'  txtDocumentosIni.Text = Format(rs!SI_DOCUMENTOS, "Standard")
'Else
'  fxFechaApertura = Format(fxFechaServidor, "dd/mm/yyyy")
'  txtEfectivoIni.Text = 0
'  txtDocumentosIni.Text = 0
'End If
'rs.Close
'
'strSQL = "select SUM(D.monto)  as monto from SIF_TRANSACCIONES_PAGO D " _
'        & " inner join SIF_FORMAS_PAGO F on D.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
'        & " and F.EFECTIVO = 1 and F.APLICA_SALDOS_FAVOR = 0" _
'        & " where cod_apertura = " & ModuloCajas.mApertura & "  and COD_CAJA = '" & ModuloCajas.mCaja & "'"
'rs.Open strSQL, glogon.Conection, adOpenStatic
'
'If Not rs.EOF Then
'    txtEfectivoFin.Text = Format(rs!Monto, "Standard")
'Else
'    txtEfectivoFin.Text = Format(0, "Standard")
'End If
'rs.Close
'
'strSQL = "select SUM(D.monto)  as monto from SIF_TRANSACCIONES_PAGO D " _
'        & " inner join SIF_FORMAS_PAGO F on D.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
'        & " and F.EFECTIVO = 0 and F.APLICA_SALDOS_FAVOR = 0" _
'        & " where cod_apertura = " & ModuloCajas.mApertura & "  and COD_CAJA = '" & ModuloCajas.mCaja & "'"
'rs.Open strSQL, glogon.Conection, adOpenStatic
'
'If Not rs.EOF Then
'    txtDocumentosFin.Text = Format(rs!Monto, "Standard")
'Else
'    txtDocumentosFin.Text = Format(0, "Standard")
'End If
'rs.Close
'
'strSQL = "select SUM(D.monto)  as monto from SIF_TRANSACCIONES_PAGO D " _
'        & " inner join SIF_FORMAS_PAGO F on D.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
'        & " and F.APLICA_SALDOS_FAVOR = 1 and F.efectivo = 0" _
'        & " where cod_apertura = " & ModuloCajas.mApertura & "  and COD_CAJA = '" & ModuloCajas.mCaja & "'"
'rs.Open strSQL, glogon.Conection, adOpenStatic
'
'If Not rs.EOF Then
'    txtSaldosFavor.Text = Format(rs!Monto, "Standard")
'Else
'    txtSaldosFavor.Text = Format(0, "Standard")
'End If
'rs.Close

End Sub

Private Sub tblAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String


On Error GoTo vError

Select Case Button.Key
   Case "Aplicar"
        If fxValida Then
          strSQL = "UPDATE CAJAS_APERTURAS_CIERRES set Estado = 'C',CIERRE_FECHA = getdate()" _
                  & ", CIERRE_USUARIO ='" & glogon.Usuario & "', DP_NUMERO = '" & txtDepNum.Text & "'" _
                  & " , DP_FECHA = getdate(), DP_MONTO_DOC =" & CCur(txtMontoDoc.Text) & " " _
                  & " , DP_MONTO_EFECTIVO =" & CCur(txtMontoEfectivo.Text) & "" _
                  & " where cod_apertura = " & ModuloCajas.mApertura & "  and COD_CAJA = '" & ModuloCajas.mCaja & "'"
          glogon.Conection.Execute strSQL
          
          Call sbImprimeDocumento
          
          MsgBox "Cierre aplicado Satisfactoriamente"
          
        End If
   Case "Cancelar"
     Unload Me
     Exit Sub
End Select

Exit Sub

vError:
   MsgBox Err.Description, vbCritical


End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If Val(txtMontoEfectivo.Text) = 0 And Val(txtMontoDoc.Text) Then vMensaje = vMensaje & vbCrLf & " - Monto del Depósito Invalido..."

If Trim(txtDepNum.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Número del Depósito Invalido..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function



Private Sub sbImprimeDocumento()
Dim strSQL As String, x As New clsImpresoras
Dim vFlat As Boolean, rs As New ADODB.Recordset
Dim vEmpresa As String, vCedJur As String

On Error GoTo vError:

strSQL = "select nombre from sif_empresa"
rs.Open strSQL, glogon.Conection, adOpenStatic
 vEmpresa = UCase(rs!Nombre & "")
rs.Close

Call sbDepositoSistema

With frmContenedor.Crt
   .Reset
   .WindowShowPrintSetupBtn = True
   .WindowState = crptMaximized
   
   .Connect = glogon.ConectRPT
   .PrinterDriver = x.Controlador
   .PrinterName = x.Nombre
   .PrinterPort = x.Puerto
   

   
   '.Destination = crptToPrinter
   .Destination = crptToWindow
   
   
   .Formulas(0) = "Empresa = '" & vEmpresa & "'"
   .Formulas(1) = "fxUsuario = '" & glogon.Usuario & "'"
   .Formulas(2) = "fxCierre = '" & ModuloCajas.mTipoCierre & "'"
   .Formulas(3) = "Efectivo = '" & Format(cMontoEfectivo, "Standard") & "'"
   .Formulas(4) = "Documentos = '" & Format(cMontoDocumentos, "Standard") & "'"
   .Formulas(5) = "Total = '" & Format(cMontoDocumentos + cMontoEfectivo, "Standard") & "'"
   .ReportFileName = SIFGlobal.fxSIFPathReportes("Cajas_Cierre.rpt")
   
   .SelectionFormula = "{CAJAS_APERTURAS_CIERRES.COD_CAJA} = '" & ModuloCajas.mCaja _
                     & "' AND {CAJAS_APERTURAS_CIERRES.COD_APERTURA} = " & ModuloCajas.mApertura & ""
    .WindowTitle = "Cierre de Caja"
    
   '.PrintReport
  .Action = 1
End With

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
End Sub

Private Sub txtMontoDoc_Change()
If Not IsNumeric(txtMontoDoc) Then
    MsgBox "Debe digitar solamente números..."
    txtMontoDoc.Text = 0
    txtMontoDoc.SetFocus
End If

End Sub

Private Sub txtMontoDoc_LostFocus()
txtMontoDoc.Text = Format(txtMontoDoc.Text, "Standard")
End Sub

Private Sub txtMontoEfectivo_Change()
If Not IsNumeric(txtMontoEfectivo) Then
    MsgBox "Debe digitar solamente números..."
    txtMontoEfectivo.Text = 0
    txtMontoEfectivo.SetFocus
End If
End Sub

Private Sub txtMontoEfectivo_LostFocus()
txtMontoEfectivo.Text = Format(txtMontoEfectivo.Text, "Standard")
End Sub



Private Sub sbDepositoSistema()
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select SUM(D.monto)  as monto from SIF_TRANSACCIONES_PAGO D " _
        & " inner join SIF_FORMAS_PAGO F on D.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
        & " and F.EFECTIVO = 1 and F.APLICA_SALDOS_FAVOR = 0 and  F.APLICA_PARA_DEPOSITO = 1" _
        & " where cod_apertura = " & ModuloCajas.mApertura & "  and COD_CAJA = '" & ModuloCajas.mCaja & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
    cMontoEfectivo = Format(rs!Monto, "Standard")
Else
    cMontoEfectivo = Format(0, "Standard")
End If
rs.Close

strSQL = "select SUM(D.monto)  as monto from SIF_TRANSACCIONES_PAGO D " _
        & " inner join SIF_FORMAS_PAGO F on D.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
        & " and F.EFECTIVO = 0 and F.APLICA_SALDOS_FAVOR = 0 and F.APLICA_PARA_DEPOSITO = 1" _
        & " where cod_apertura = " & ModuloCajas.mApertura & "  and COD_CAJA = '" & ModuloCajas.mCaja & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
    cMontoDocumentos = Format(rs!Monto, "Standard")
Else
    cMontoDocumentos = Format(0, "Standard")
End If
rs.Close

End Sub
