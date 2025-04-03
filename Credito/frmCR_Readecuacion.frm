VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCR_CorreccionCreditos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Corrección de Créditos Activos y Retenciones"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   HelpContextID   =   3013
   Icon            =   "frmCR_Readecuacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Documento"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   10
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   4200
      Width           =   2055
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Fiadores"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   9
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3900
      Width           =   2055
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abonos Especiales"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   8
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3600
      Width           =   2055
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio Primer Deducción"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   7
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3300
      Width           =   2055
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio Ult. Abono"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   6
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3000
      Width           =   2055
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   1410
      Left            =   0
      TabIndex        =   29
      Top             =   4920
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   2487
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha Pro"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Fecha Sist."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Int.Cor."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Int.Mor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amortización"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Elimina Cuotas en Mora >>"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   5
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2700
      Width           =   2055
   End
   Begin VB.Frame fraOpcion 
      Caption         =   "Cambios"
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   4680
      TabIndex        =   19
      Top             =   1200
      Width           =   2535
      Begin VB.TextBox txtAmortizacion 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   44
         ToolTipText     =   "Número de Operación"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtIntereses 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   43
         ToolTipText     =   "Número de Operación"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkPlazoSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Utilizar Saldo (como Base)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   2295
      End
      Begin VB.VScrollBar vsBar 
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   840
         Width           =   270
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortización"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   41
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   2400
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lblOpcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   2400
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Cuota"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   4
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado Actual"
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   2535
      Begin VB.Label lblPrideduc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   39
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1° Deduc"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label txtUltMov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   34
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblCuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblPlazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Interés"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plazo"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Monto "
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   3
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2100
      Width           =   2055
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Código"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Interés"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1500
      Width           =   2055
   End
   Begin VB.OptionButton optCorreccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambio de Plazo"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtOperacion 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Número de Operación"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblOpex 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   6600
      TabIndex        =   36
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccione las Cuotas Morosas para Anulación y Luego Presione Aceptar"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   4680
      Width           =   7335
   End
   Begin VB.Image imgExcluirCredito 
      Height          =   255
      Left            =   6960
      Picture         =   "frmCR_Readecuacion.frx":030A
      Stretch         =   -1  'True
      ToolTipText     =   "Excluye Esta operacion por Medio de Anulacion en la B.D."
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "# Operación"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Cédula"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Descripción del Código"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Nombre de la Persona"
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblCodigo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      ToolTipText     =   "Código del Préstamo"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblCedula 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Cédula de la Persona"
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmCR_CorreccionCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vRetencion As Boolean

Private Function fxValidaCambio() As Boolean
Dim rs As New ADODB.Recordset

fxValidaCambio = True

If optCorreccion(2).Value = True Then
    rs.CursorLocation = adUseServer
    rs.Open "select coalesce(count(*),0) as Existe from catalogo where codigo = '" & txtCambio _
            & "'", glogon.Conection, adOpenStatic
    fxValidaCambio = IIf((rs!existe = 1), True, False)
    txtCambio = UCase(txtCambio)
    rs.Close
Else
  If IsNumeric(txtCambio) Then
    If optCorreccion(0).Value = True Or optCorreccion(1).Value = True Then
      txtCambio = CLng(txtCambio)
    Else
      txtCambio = CCur(txtCambio)
    End If
    
    If optCorreccion(1).Value = True And txtCambio > 99 Then fxValidaCambio = False
  Else
    fxValidaCambio = False
  End If
End If

If vOperacion = 0 Then fxValidaCambio = False

End Function


Private Sub sbEliminaMora()
Dim strSQL As String, itmX As ListItem, lng As Long
Dim i As Integer

i = 0

For lng = 1 To lsw.ListItems.Count
Set itmX = lsw.FindItem(lng, lvwTag)
 If itmX Is Nothing Then  'No lo encontro
  'nada
 Else
  lsw.ListItems.Item(lng).Selected = lng
  If lsw.SelectedItem.Checked Then
    strSQL = "update morosidad set estado = 'N' where id_moro = " & lsw.SelectedItem.Text
    glogon.Conection.Execute strSQL
    i = i + 1
    Call sbBitacoraCredito("06", "ID:" & lsw.SelectedItem.Text, "C", txtOperacion, lblCodigo.Caption)
    Call Bitacora("Anula", "Morosidad OP: " & txtOperacion & " ID:" & lsw.SelectedItem.Text)
  End If
 End If
Next lng

 MsgBox "Anulaciones Realizadas Satisfactoriamente..."
 
optCorreccion(1).Value = True
Call sbFormSize

End Sub


Private Sub sbAbonoEspecial()
Dim iRespuesta As Integer, strSQL As String
Dim vOB As String, lngRecibo As Long, vCuenta As String
Dim rs As New ADODB.Recordset

'Verificar que la amortizacion no sea mayor que el saldo y
'que no tenga cuotas morosas
'No tiene que afectar el ultimo movimiento

If vOperacion = 0 Then
    MsgBox "Ingrese un número de operacion válido...", vbCritical
    Exit Sub
End If

If CCur(txtAmortizacion) > CCur(lblSaldo.Caption) Then
    MsgBox "La Amortización Especificada Es Mayor al Saldo, verifique...", vbCritical
    Exit Sub
End If

strSQL = "select count(*) as Total from morosidad where estado = 'A' and id_solicitud = " & vOperacion
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs!total > 0 Then
  MsgBox "No se puede Aplicar Abono Especial porque esta operación se encuentra en mora...", vbCritical
  Exit Sub
End If
rs.Close

vOB = ""

iRespuesta = MsgBox("Esta seguro de realizar abono especial a la OP " & vOperacion, vbYesNo)

vCuenta = Trim(fxDocumentoCuenta("NC"))

If vAseDocValido = False Then
    Me.MousePointer = vbDefault
    MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
          & " válida para esta operación...", vbCritical
    Exit Sub
End If

lngRecibo = 0
If uRecibos Then lngRecibo = fxDocumentoAbono5(CCur(txtAmortizacion), CCur(txtIntereses), 0, "NC", vCuenta)

If iRespuesta = 6 Then
  vOB = "- ABONO ESPECIAL NC:" & lngRecibo
  
  strSQL = "Update reg_creditos set estado = '" & IIf((CCur(txtAmortizacion) >= CCur(lblSaldo.Caption)), "C", "A") & "'," _
         & "SALDO = SALDO - " & CCur(txtAmortizacion) & ",AMORTIZA=AMORTIZA + " & CCur(txtAmortizacion) _
         & ",interesc = interesc + " & CCur(txtIntereses) _
         & " where id_solicitud = " & vOperacion
  glogon.Conection.Execute strSQL
  
  strSQL = "INSERT CREDITOS_DT(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS,FECHAP,TCON,NCON,ESTADO)" _
         & "VALUES('" & lblCodigo.Caption & "'," & vOperacion & ",0," & CCur(txtAmortizacion) + CCur(txtIntereses) _
         & "," & CCur(txtIntereses) & "," & CCur(txtAmortizacion) & ",'" & Format(fxFechaServidor, "yyyy/mm/dd") _
         & "'," & GLOBALES.glngFechaCR & ",7," & IIf((lngRecibo = 0), "null", lngRecibo) & ",'A')"
  glogon.Conection.Execute strSQL
  
  Call Bitacora("Aplica", "Abono Especial de la operación :" & vOperacion)
  
  If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "NC")
  
  MsgBox "Abono Especial aplicado con Nota de Crédito # " & lngRecibo, vbInformation
  
  Call sbCargaOperacion
  
End If

End Sub

Private Sub cmdAceptar_Click()
Dim iRespuesta As Integer, strSQL As String, vDH As String
Dim strDetalleBitacora As String, vCuenta As String, curMonto As Currency
Dim lngRecibo As Long, vTipo As String

 txtCambio.Locked = False
 chkPlazoSaldo.Enabled = False
 vsBar.Enabled = False

Select Case True
  Case Me.optCorreccion(5).Value
      Call sbEliminaMora
      Exit Sub
  Case Me.optCorreccion(8).Value
      Call sbAbonoEspecial
      Exit Sub
  Case Me.optCorreccion(9).Value 'Cambio de Fiadores
      Operacion.Operacion = txtOperacion
      frmCR_SolicitudesFiadores.Show
      Exit Sub
End Select


If Not fxValidaCambio Then
  MsgBox "Error : El valor de cambio no es válido...", vbCritical
  Exit Sub
End If
strDetalleBitacora = ""

iRespuesta = MsgBox("Esta seguro que desea " & lblOpcion.Caption & " de la OP " & vOperacion, vbYesNo)

If iRespuesta = vbNo Or iRespuesta = vbCancel Then Exit Sub
 
 Select Case True
   
   Case optCorreccion(0) 'Plazo
        
     vTipo = "OT"
     
     If chkPlazoSaldo.Value = 0 Then
        strSQL = "update reg_Creditos set plazo = " & txtCambio & ",cuota = " _
               & CCur(fxCalcula_Cuota(CCur(lblMonto.Caption), txtCambio, lblInteres.Caption)) _
               & " where id_solicitud = " & vOperacion
        strDetalleBitacora = "REF: MONTO "
     Else
        strSQL = "update reg_Creditos set plazo = " & txtCambio & ",cuota = " _
               & CCur(fxCalcula_Cuota(CCur(lblSaldo.Caption), txtCambio, lblInteres.Caption)) _
               & " where id_solicitud = " & vOperacion
        strDetalleBitacora = "REF: SALDO "
     End If
     glogon.Conection.Execute strSQL
     
     Call sbBitacoraCredito("01", ("De: " & lblPlazo.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
     
     strDetalleBitacora = strDetalleBitacora & "Cambia Plazo De " & lblPlazo.Caption & " a " & txtCambio & " OP " & vOperacion
     
   Case optCorreccion(1) 'Interes
     
     
     strSQL = "update reg_Creditos set interesv = " & txtCambio & ",cuota = " _
            & CCur(fxCalcula_Cuota(CCur(lblMonto.Caption), lblPlazo.Caption, txtCambio)) _
            & " where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
     
     Call sbBitacoraCredito("02", ("De: " & lblInteres.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
     
     strDetalleBitacora = "Cambia Interes De " & lblInteres.Caption & " a " & txtCambio & " OP " & vOperacion
     
     
   Case optCorreccion(2) 'Codigo
     
        vTipo = "ND"
        
        vCuenta = Trim(fxDocumentoCuenta("ND"))
        
        If vAseDocValido = False Then
            Me.MousePointer = vbDefault
            MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
                  & " válida para esta operación...", vbCritical
            Exit Sub
        End If
        
        lngRecibo = 0
     
     Call sbBitacoraCredito("03", ("De: " & lblCodigo.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
     
     strDetalleBitacora = "Cambia Código De " & lblCodigo.Caption & " a " & txtCambio & " OP " & vOperacion
     
     strSQL = "update creditos_dt set codigo = '" & txtCambio & "' where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
    
     strSQL = "update morosidad set codigo = '" & txtCambio & "' where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
    
     strSQL = "update refundiciones set codigo = '" & txtCambio & "' where id_solicitudr = " & vOperacion
     glogon.Conection.Execute strSQL
    
     strSQL = "update desembolsos set codigo = '" & txtCambio & "' where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
    
     strSQL = "update cobro_avisos set codigo = '" & txtCambio & "' where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
    
     strSQL = "update reg_Creditos set codigo = '" & txtCambio & "' where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
    
     If uRecibos Then lngRecibo = fxDocumentoAbono3(CCur(lblSaldo.Caption), "ND")
    
    
   Case optCorreccion(3) 'Monto
     
     If CCur(lblMonto.Caption) = CCur(txtCambio) Then Exit Sub
       
     If CCur(lblMonto.Caption) < CCur(txtCambio) Then
        curMonto = CCur(txtCambio) - CCur(lblMonto.Caption)
        vDH = "D"
     Else
        curMonto = CCur(txtCambio) - CCur(lblMonto.Caption)
        curMonto = Abs(curMonto)
        vDH = "H"
     End If
     
    
     If CCur(lblMonto.Caption) < CCur(txtCambio) Then 'ND : AUMENTA EL MONTO
            vTipo = "ND"
            vCuenta = Trim(fxDocumentoCuenta("ND"))
            
            If vAseDocValido = False Then
                Me.MousePointer = vbDefault
                MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
                      & " válida para esta operación...", vbCritical
                Exit Sub
            End If
            
            lngRecibo = 0
            
            strSQL = "update reg_Creditos set montoapr = " & txtCambio & ",cuota = " _
                   & CCur(fxCalcula_Cuota(txtCambio, lblPlazo.Caption, lblInteres.Caption)) _
                   & ",saldo = " & txtCambio & " - AMORTIZA" _
                   & " where id_solicitud = " & vOperacion
            glogon.Conection.Execute strSQL
     
     
            If uRecibos Then lngRecibo = fxDocumentoAbono4(curMonto, "ND", vCuenta, vDH)
            
            strDetalleBitacora = "Cambia Monto De:" & CCur(lblMonto.Caption) & " A:" _
                               & txtCambio & " OP:" & vOperacion & "-ND:" & lngRecibo
            
            strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
                   & "fechas,fechap,tcon,ncon,estado) values('" & lblCodigo.Caption _
                   & "'," & txtOperacion & ",0,0,0," & CCur(txtCambio) - CCur(lblMonto.Caption) & ",'" _
                   & Format(fxFechaServidor, "yyyy/mm/dd") & "'," & GLOBALES.glngFechaCR _
                   & ",8," & lngRecibo & ",'A')"
            glogon.Conection.Execute strSQL
    
     Else 'NC : DISMINUYE EL MONTO
            
            vTipo = "NC"
            vCuenta = Trim(fxDocumentoCuenta("NC"))
            
            If vAseDocValido = False Then
                Me.MousePointer = vbDefault
                MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
                      & " válida para esta operación...", vbCritical
                Exit Sub
            End If
            
            lngRecibo = 0
            
            strSQL = "update reg_Creditos set montoapr = " & CCur(txtCambio) & ",cuota = " _
                   & CCur(fxCalcula_Cuota(txtCambio, lblPlazo.Caption, lblInteres.Caption)) _
                   & ",saldo = " & CCur(txtCambio) & "-AMORTIZA" _
                   & " where id_solicitud = " & vOperacion
            glogon.Conection.Execute strSQL
            
            
            If uRecibos Then lngRecibo = fxDocumentoAbono4(curMonto, "NC", vCuenta, vDH)
            
            strDetalleBitacora = "Cambia Monto De:" & CCur(lblMonto.Caption) & " A:" _
                               & txtCambio & " OP:" & vOperacion & "-NC:" & lngRecibo
            strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
                   & "fechas,fechap,tcon,ncon,estado) values('" & lblCodigo.Caption _
                   & "'," & txtOperacion & ",0,0,0," & CCur(lblMonto.Caption) - CCur(txtCambio) & ",'" _
                   & Format(fxFechaServidor, "yyyy/mm/dd") & "'," & GLOBALES.glngFechaCR _
                   & ",7," & lngRecibo & ",'A')"
            glogon.Conection.Execute strSQL
      
      
     End If
   
        Call sbBitacoraCredito("09", ("De: " & lblMonto.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)

  
   Case optCorreccion(4) 'Cuota
     
     vTipo = "OT"
     If vRetencion Then
        strSQL = "update reg_Creditos set cuota = " & txtCambio _
               & ",montoapr = " & txtCambio & ",saldo = " & txtCambio _
               & " where id_solicitud = " & vOperacion
     Else
        strSQL = "update reg_Creditos set cuota = " & txtCambio _
               & " where id_solicitud = " & vOperacion
     End If
     glogon.Conection.Execute strSQL
 
     Call sbBitacoraCredito("03", ("De: " & CCur(lblCuota.Caption) & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
 
     strDetalleBitacora = "Cambia Cuota De " & CCur(lblCuota.Caption) & " a " & txtCambio & " OP " & vOperacion
 
   Case optCorreccion(6) 'Ultimo Abono
      
     vTipo = "OT"
      
     vsBar.Enabled = False
     strSQL = "update reg_Creditos set fecult = " & txtCambio _
            & " where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
 

     Call sbBitacoraCredito("04", ("De: " & txtUltMov & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
     
     strDetalleBitacora = "Fecha Ult. Abono DE " & txtUltMov & " A " & txtCambio & " OP " & vOperacion
 
   Case optCorreccion(7) 'Primer Deducción
      
     vTipo = "OT"
     
     lngRecibo = 0
      
      vsBar.Enabled = False
     strSQL = "update reg_Creditos set prideduc = " & txtCambio _
            & " where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
 
     Call sbBitacoraCredito("05", ("De: " & lblPrideduc.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
 
     strDetalleBitacora = "Fecha 1er Deducción DE " & lblPrideduc.Caption & " A " & txtCambio & " OP " & vOperacion
 
 
 End Select
 
 Call Bitacora("Aplica", strDetalleBitacora)
 
 Call sbCargaOperacion
 
 txtCambio = ""
  
Select Case vTipo
  Case "ND"
    MsgBox "Cambio Realizado Satisfactoriamente..." _
      & vbCrLf & "Se Generó Nota de Débito # " & lngRecibo, vbInformation
      If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "ND")
  Case "NC"
    MsgBox "Cambio Realizado Satisfactoriamente..." _
      & vbCrLf & "Se Generó Nota de Crédito # " & lngRecibo, vbInformation
     If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "NC")
  Case "OT"
    MsgBox "Cambio Realizado Satisfactoriamente..."
End Select

End Sub

Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub sbFormSize()
If optCorreccion.Item(5).Value Then
  Me.Height = 6735
Else
  Me.Height = 4980
End If

End Sub

Private Sub sbCargaOperacion()
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer

vOperacion = 0

On Error Resume Next

strSQL = "select R.*,S.nombre,C.descripcion,C.retencion,C.poliza" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula " _
       & " where R.estado = 'A' and R.proceso = 'N'" _
       & " and id_solicitud =" & txtOperacion
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF And Not rs.BOF Then
 lblCedula.Caption = rs!Cedula
 lblCodigo.Caption = rs!Codigo
 lblNombre.Caption = rs!Nombre
 lblDescripcion.Caption = rs!Descripcion
 lblOpex.Caption = IIf((rs!opex = 1), "OPEX", "")

 lblMonto.Caption = Format(rs!montoapr, "###,###,###,##0.00")
 lblSaldo.Caption = Format(rs!Saldo, "###,###,###,##0.00")
 lblCuota.Caption = Format(rs!cuota, "###,###,###,##0.00")
 txtUltMov = IIf(IsNull(rs!fecult), 0, rs!fecult)
 lblPrideduc.Caption = IIf(IsNull(rs!prideduc), 0, rs!prideduc)
 lblPlazo.Caption = rs!Plazo
 lblInteres.Caption = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
 vOperacion = rs!id_solicitud
 
 If rs!retencion = "S" Or rs!poliza = "S" Then
   vRetencion = True
   
   optCorreccion(4).Enabled = True
   optCorreccion(5).Enabled = True
   optCorreccion(6).Enabled = True
   optCorreccion(7).Enabled = True
   optCorreccion(8).Enabled = True
   
   
 Else 'No es Retencion aplican todas las funciones
 
   vRetencion = False
   For i = 0 To 10
     optCorreccion(i).Enabled = True
   Next i
 End If
 
 
Else
 vOperacion = 0
 MsgBox "La operación no se encontró o está cancelada, o Pertenece a un codigo de Retencion...", vbInformation
End If

rs.Close

End Sub

Private Sub Form_Load()
vModulo = 3
Call Formularios(Me)

Call txtOperacion_Change
Call RefrescaTags(Me)

sbFormSize
fraOpcion.Enabled = False
imgExcluirCredito.Visible = IIf((cmdAceptar.Tag = 1), True, False)

End Sub



Private Sub imgExcluirCredito_Click()
Dim iRespuesta As Integer, strSQL As String
Dim vOB As String, lngRecibo As Long, vCuenta As String

If vOperacion = 0 Then
    MsgBox "Ingrese un número de operacion válido...", vbInformation
    Exit Sub
End If

vOB = ""

iRespuesta = MsgBox("Esta seguro que desea Excluir del Sistema la OP " & vOperacion, vbYesNo)

lngRecibo = 0

If iRespuesta = vbYes Then
  
  If Not vRetencion Then
        
        vCuenta = Trim(fxDocumentoCuenta("NC"))
        
        If vAseDocValido = False Then
            Me.MousePointer = vbDefault
            MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
                  & " válida para esta operación...", vbCritical
            Exit Sub
        End If
        
        If uRecibos Then lngRecibo = fxDocumentoAbono(CCur(lblSaldo.Caption), "NC", vCuenta)
        
        vOB = "SE EXLUYE CON NOTA CREDITO # " & lngRecibo
        
        strSQL = "Update reg_creditos set estado = 'C',SALDO=0,AMORTIZA=MONTOAPR" _
               & ", observacion = observacion + '" & vOB & "'" _
               & " where id_solicitud = " & vOperacion
        glogon.Conection.Execute strSQL
        
        strSQL = "INSERT CREDITOS_DT(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS,FECHAP,TCON,NCON,ESTADO)" _
               & "VALUES('" & lblCodigo.Caption & "'," & vOperacion & ",0," & CCur(lblSaldo.Caption) & ",0," & CCur(lblSaldo.Caption) _
               & ",'" & Format(fxFechaServidor, "yyyy/mm/dd") & "'," & GLOBALES.glngFechaCR & ",7," _
               & IIf((lngRecibo = 0), "null", lngRecibo) & ",'A')"
        glogon.Conection.Execute strSQL
        
        Call Bitacora("Aplica", "Exclusión de la operación :" & vOperacion)
        
        If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "NC")
        
        MsgBox "Exclusión aplicada con Nota de Crédito # " & lngRecibo, vbInformation
        vOperacion = 0
  
  Else 'Es una Retencion
        
        strSQL = "Update reg_creditos set estado = 'C',SALDO=0" _
               & " where id_solicitud = " & vOperacion
        glogon.Conection.Execute strSQL
        
        Call sbBitacoraCredito("07", "Monto :" & lblSaldo.Caption, IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
        
        Call Bitacora("Aplica", "Exclusión de la operación :" & vOperacion)
        
        MsgBox "Exclusión aplicada y Guardada en Bitacora Creditos...", vbInformation
        vOperacion = 0
  
  End If 'Retencion

End If

End Sub

Private Function fxDocumentoAbono(curAmortiza As Currency, vTipoDoc As String, vCuenta As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, rs2 As New ADODB.Recordset, strCliente As String

lngRecibo = fxDocumentoConsecutivo(vTipoDoc)
fxDocumentoAbono = lngRecibo

If UCase(lblOpex.Caption) = "OPEX" Then
  strSQL = "select ctaoamort as ctaAmortiza,ctaointc as ctaintc,ctaointm as ctaintm from catalogo"
Else
  strSQL = "select ctanamort as ctaAmortiza,ctanintc as ctaintc,ctanintm as ctaintm from catalogo"
End If
  strSQL = strSQL & " where codigo = '" & lblCodigo.Caption & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

rs2.CursorLocation = adUseServer
rs2.Open "select saldo from reg_creditos where id_solicitud = " & txtOperacion, glogon.Conection, adOpenStatic

strLinea(1) = "Saldo Anterior    " & Format(rs2!Saldo, "###,###,###,##0.00")
strLinea(2) = "Interes Corriente " & "0.00"
strLinea(3) = "Interes Moratorio " & "0.00"
strLinea(4) = "Amortizacion      " & Format(curAmortiza, "###,###,###,##0.00")
strLinea(5) = "Saldo Actual      " & Format(rs2!Saldo - curAmortiza, "###,###,###,##0.00")
strLinea(6) = "Operación         " & txtOperacion
strLinea(7) = "Código            " & lblCodigo.Caption & "-" & UCase(lblOpex.Caption)
strLinea(8) = "Descripción       " & fxDescribeCodigo(lblCodigo.Caption)
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = "EXCLUYE   "

rs2.Close

strCliente = Trim(lblCedula) & " - " & Trim(lblNombre)
strCliente = Mid(strCliente, 1, 45)


strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,DP)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") & "','" & strCliente & "','" _
        & "EXCLUYE LA Op:" & txtOperacion & "'," & curAmortiza & ",'" & glogon.Usuario & "','P','" _
        & "N','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"

glogon.Conection.Execute strSQL

'ASIENTO

If curAmortiza > 0 Then
  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & curAmortiza & ",'H')"
  glogon.Conection.Execute strSQL
End If

If curAmortiza > 0 Then
  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & vCuenta & "'," & curAmortiza & ",'D')"
  glogon.Conection.Execute strSQL
End If

rs.Close


End Function


Private Function fxDocumentoAbono2(vTipoDoc As String, vCuenta As String, vDetalle As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, strCliente As String

lngRecibo = fxDocumentoConsecutivo(vTipoDoc)
fxDocumentoAbono2 = lngRecibo

strLinea(1) = lblOpcion.Caption
strLinea(2) = vDetalle
strLinea(3) = ""
strLinea(4) = ""
strLinea(5) = ""
strLinea(6) = "Operación         " & txtOperacion
strLinea(7) = "Código            " & lblCodigo.Caption & "-" & UCase(lblOpex.Caption)
strLinea(8) = "Descripción       " & fxDescribeCodigo(lblCodigo.Caption)
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = ""

strCliente = Trim(lblCedula) & " - " & Trim(lblNombre)
strCliente = Mid(strCliente, 1, 45)


strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,DP)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") & "','" & strCliente & "','" _
        & UCase(lblOpcion.Caption) & "- OP " & txtOperacion & "',0,'" & glogon.Usuario & "','P','" _
        & "N','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"

glogon.Conection.Execute strSQL

'ASIENTO

strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & vCuenta & "',1,'H')"
glogon.Conection.Execute strSQL

strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & vCuenta & "',1,'D')"
glogon.Conection.Execute strSQL


End Function


Private Function fxDocumentoAbono3(curAmortiza As Currency, vTipoDoc As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, rs2 As New ADODB.Recordset, strCliente As String

lngRecibo = fxDocumentoConsecutivo(vTipoDoc)
fxDocumentoAbono3 = lngRecibo

If UCase(lblOpex.Caption) = "OPEX" Then
  strSQL = "select ctaoamort as ctaAmortiza,ctaointc as ctaintc,ctaointm as ctaintm from catalogo"
Else
  strSQL = "select ctanamort as ctaAmortiza,ctanintc as ctaintc,ctanintm as ctaintm from catalogo"
End If
  strSQL = strSQL & " where codigo = '" & lblCodigo.Caption & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

rs2.CursorLocation = adUseServer
rs2.Open "select saldo from reg_creditos where id_solicitud = " & txtOperacion, glogon.Conection, adOpenStatic

strLinea(1) = "CAMBIO DE CODIGO"
strLinea(2) = "DE " & lblCodigo.Caption & "-" & UCase(lblOpex.Caption)
strLinea(3) = "A " & txtCambio & "-" & UCase(lblOpex.Caption)
strLinea(4) = ""
strLinea(5) = ""
strLinea(6) = "Operación         " & txtOperacion
strLinea(7) = "Código            " & lblCodigo.Caption & "-" & UCase(lblOpex.Caption)
strLinea(8) = "Descripción       " & fxDescribeCodigo(lblCodigo.Caption)
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = ""

rs2.Close

If UCase(lblOpex.Caption) = "OPEX" Then
  strSQL = "select ctaoamort as ctaAmortiza,ctaointc as ctaintc,ctaointm as ctaintm from catalogo"
Else
  strSQL = "select ctanamort as ctaAmortiza,ctanintc as ctaintc,ctanintm as ctaintm from catalogo"
End If
  strSQL = strSQL & " where codigo = '" & txtCambio & "'"
rs2.Open strSQL, glogon.Conection, adOpenStatic

strCliente = Trim(lblCedula) & " - " & Trim(lblNombre)
strCliente = Mid(strCliente, 1, 45)


strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,DP)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") & "','" & strCliente & "','" _
        & "CAMBIO DE CODIGO OP" & txtOperacion & "',0,'" & glogon.Usuario & "','P','" _
        & "N','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"

glogon.Conection.Execute strSQL

'ASIENTO

strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs2!ctaamortiza) & "'," & curAmortiza & ",'D')"
glogon.Conection.Execute strSQL

strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & curAmortiza & ",'H')"
glogon.Conection.Execute strSQL

rs.Close
rs2.Close

End Function


Private Function fxDocumentoAbono4(curAmortiza As Currency, vTipoDoc As String, vCuenta, vDH As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, rs2 As New ADODB.Recordset, strCliente As String

lngRecibo = fxDocumentoConsecutivo(vTipoDoc)
fxDocumentoAbono4 = lngRecibo

If UCase(lblOpex.Caption) = "OPEX" Then
  strSQL = "select ctaoamort as ctaAmortiza,ctaointc as ctaintc,ctaointm as ctaintm from catalogo"
Else
  strSQL = "select ctanamort as ctaAmortiza,ctanintc as ctaintc,ctanintm as ctaintm from catalogo"
End If
  strSQL = strSQL & " where codigo = '" & lblCodigo.Caption & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

rs2.CursorLocation = adUseServer
rs2.Open "select saldo from reg_creditos where id_solicitud = " & txtOperacion, glogon.Conection, adOpenStatic

strLinea(1) = "CAMBIO DE MONTO"
strLinea(2) = "DE " & CCur(lblMonto.Caption)
strLinea(3) = "A " & CCur(txtCambio)
strLinea(4) = ""
strLinea(5) = ""
strLinea(6) = "Operación         " & txtOperacion
strLinea(7) = "Código            " & lblCodigo.Caption & "-" & UCase(lblOpex.Caption)
strLinea(8) = "Descripción       " & fxDescribeCodigo(lblCodigo.Caption)
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = ""

rs2.Close

strCliente = Trim(lblCedula) & " - " & Trim(lblNombre)
strCliente = Mid(strCliente, 1, 45)


strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,dp)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") & "','" & strCliente & "','" _
        & "CAMBIO DE MONTO OP" & txtOperacion & "',0,'" & glogon.Usuario & "','P','" _
        & "N','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"

glogon.Conection.Execute strSQL

'ASIENTO

strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & curAmortiza & ",'" & vDH & "')"
glogon.Conection.Execute strSQL

strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & vCuenta & "'," & curAmortiza _
        & ",'" & IIf((vDH = "D"), "H", "D") & "')"
glogon.Conection.Execute strSQL

rs.Close

End Function

Private Function fxDocumentoAbono5(curAmortiza As Currency, curIntCor As Currency, curIntMor As Currency, vTipoDoc As String, vCuenta As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, rs2 As New ADODB.Recordset, strCliente As String

lngRecibo = fxDocumentoConsecutivo(vTipoDoc)
fxDocumentoAbono5 = lngRecibo

If UCase(lblOpex.Caption) = "OPEX" Then
  strSQL = "select ctaoamort as ctaAmortiza,ctaointc as ctaintc,ctaointm as ctaintm from catalogo"
Else
  strSQL = "select ctanamort as ctaAmortiza,ctanintc as ctaintc,ctanintm as ctaintm from catalogo"
End If
  strSQL = strSQL & " where codigo = '" & lblCodigo.Caption & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

strLinea(1) = "Saldo Anterior    " & Format(CCur(lblSaldo.Caption), "###,###,###,##0.00")
strLinea(2) = "Interes Corriente " & Format(curIntCor, "###,###,###,##0.00")
strLinea(3) = "Interes Moratorio " & Format(curIntMor, "###,###,###,##0.00")
strLinea(4) = "Amortizacion      " & Format(curAmortiza, "###,###,###,##0.00")
strLinea(5) = "Saldo Actual      " & Format(CCur(lblSaldo.Caption) - curAmortiza, "###,###,###,##0.00")
strLinea(6) = "Operación         " & txtOperacion
strLinea(7) = "Código            " & lblCodigo.Caption & "-" & UCase(lblOpex.Caption)
strLinea(8) = "Descripción       " & fxDescribeCodigo(lblCodigo.Caption)
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = "Abono Especial   "


strCliente = Trim(lblCedula) & " - " & Trim(lblNombre)
strCliente = Mid(strCliente, 1, 45)


strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,dp)" _
        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") & "','" & strCliente & "','" _
        & "ABONO ESPECIAL OP :" & txtOperacion & "'," & curAmortiza & ",'" & glogon.Usuario & "','P','" _
        & "N','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"

glogon.Conection.Execute strSQL

'ASIENTO

If curAmortiza > 0 Then
  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & curAmortiza & ",'H')"
  glogon.Conection.Execute strSQL
End If


If curIntCor > 0 Then
  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaintc) & "'," & curIntCor & ",'H')"
  glogon.Conection.Execute strSQL
End If

If curIntMor > 0 Then
  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaintm) & "'," & curIntMor & ",'H')"
  glogon.Conection.Execute strSQL
End If


If curAmortiza + curIntCor + curIntMor > 0 Then
  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & vCuenta & "'," & curAmortiza + curIntCor + curIntMor & ",'D')"
  glogon.Conection.Execute strSQL
End If

rs.Close


End Function



Private Sub sbLlenaMorosidad()
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListItem

On Error Resume Next

If txtOperacion = "" Then Exit Sub
If Not IsNumeric(txtOperacion) Then Exit Sub

lsw.ListItems.Clear
rs.CursorLocation = adUseServer
strSQL = "select * from morosidad where estado = 'A' and id_solicitud = " & txtOperacion _
       & " Order by fechap desc"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(lsw.ListItems.Count + 1, , rs!id_moro)
     itmX.Tag = itmX.Index
     itmX.SubItems(1) = Format(rs!fechap, "####-##")
     itmX.SubItems(2) = Format(rs!fecult, "dd/mm/yyyy")
     itmX.SubItems(3) = Format(rs!intc, "###,###,###,##0.00")
     itmX.SubItems(4) = Format(rs!intm, "###,###,###,##0.00")
     itmX.SubItems(5) = Format(rs!Amortiza, "###,###,###,##0.00")
 rs.MoveNext
Loop
rs.Close
End Sub


Private Sub optCorreccion_Click(Index As Integer)
fraOpcion.Enabled = True
lblOpcion.Caption = optCorreccion(Index).Caption
txtCambio.Enabled = True
txtCambio = ""
txtCambio.SetFocus


txtIntereses = 0
txtAmortizacion = 0
txtIntereses.Enabled = False
txtAmortizacion.Enabled = False

Select Case True
  Case optCorreccion.Item(0).Value
    chkPlazoSaldo.Enabled = True
    chkPlazoSaldo.Value = 0
  Case optCorreccion.Item(5).Value
    Call sbLlenaMorosidad
  Case optCorreccion.Item(6).Value
    txtCambio = txtUltMov.Caption
    vsBar.Value = 1000
    vsBar.Tag = vsBar.Value
    vsBar.Enabled = True
    txtCambio.Locked = True
  Case optCorreccion.Item(7).Value
    txtCambio = lblPrideduc.Caption
    vsBar.Value = 1000
    vsBar.Tag = vsBar.Value
    vsBar.Enabled = True
    txtCambio.Locked = True
  Case optCorreccion.Item(8).Value 'Abonos Especiales
    txtIntereses.Enabled = True
    txtAmortizacion.Enabled = True
    txtCambio.Enabled = False
    txtIntereses.SetFocus
  Case Else
    txtCambio.Locked = False
    chkPlazoSaldo.Enabled = False
    chkPlazoSaldo.Value = 0
    vsBar.Enabled = False
End Select

Call sbFormSize

End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then cmdAceptar.SetFocus
End Sub

Private Sub txtOperacion_Change()
Dim i As Integer

 vOperacion = 0
 lblCedula.Caption = ""
 lblCodigo.Caption = ""
 lblNombre.Caption = ""
 lblDescripcion.Caption = ""
 txtUltMov = 0
 lblMonto.Caption = Format(0, "Standard")
 lblSaldo.Caption = Format(0, "Standard")
 lblCuota.Caption = Format(0, "Standard")
 lblOpex.Caption = ""
 
 lblPlazo.Caption = 0
 lblInteres.Caption = 0
 txtCambio = ""
 
 For i = 0 To 10
   optCorreccion(i).Enabled = False
 Next i
 
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then Call sbCargaOperacion
End Sub


Private Sub vsBar_Change()
On Error Resume Next
If vsBar.Value < Val(vsBar.Tag) Then txtCambio = fxFechaProcesoSiguiente(txtCambio)
If vsBar.Value > Val(vsBar.Tag) Then txtCambio = fxFechaProcesoAnterior(txtCambio)

vsBar.Tag = vsBar.Value

End Sub



