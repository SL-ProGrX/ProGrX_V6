VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInsPolizaAjustes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambios / Ajustes a Pólizas"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInsPolizaAjustes.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   2640
      Width           =   4575
   End
   Begin VB.TextBox txtPlazo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   22
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox txtCuota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1920
      TabIndex        =   21
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   20
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Picture         =   "frmInsPolizaAjustes.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdAjustar 
      Caption         =   "&Ajustar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      Picture         =   "frmInsPolizaAjustes.frx":6A39
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtTipoCuentaCod 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtTipoSeguroCod 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtTipoCuentaDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   2160
      Width           =   4935
   End
   Begin VB.TextBox txtTipoSeguroDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   1800
      Width           =   4935
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   1080
      Width           =   5535
   End
   Begin VB.TextBox txtVendedorCod 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtVendedorDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Presiones F4 para Consultar"
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox txtPoliza 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoSeguro 
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   1800
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1572865
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoCuenta 
      Height          =   255
      Left            =   8640
      TabIndex        =   12
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1572865
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":6BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":6CFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":6E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":6F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":7051
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":716F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":7299
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":73BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":74DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":75D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsPolizaAjustes.frx":76EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   5685
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario Activa"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha Activa"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario - Cierra"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha Cierre"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4080
      TabIndex        =   30
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1572865
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3840
      TabIndex        =   28
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   25
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   24
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblPlazo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblPagador 
      Caption         =   "Tipo Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblContrato 
      Caption         =   "Tipo Seguro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   9480
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   9480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image ImgAutorizacion 
      Height          =   255
      Left            =   4680
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "No. Poliza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblNombre 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmInsPolizaAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean
Dim vFecha          As Date

Function fxPersonaNombre(strCedula As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select nombre from Socios where cedula = '" & strCedula & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxPersonaNombre = ""
Else
 fxPersonaNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close

End Function




Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

fxValida = True
vMensaje = ""


If Len(txtPoliza.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- No se indicó el número de la póliza!"


If IsNumeric(txtPlazo) Then
 If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo es Inválido"
End If


If IsNumeric(txtMonto.Text) Then
 If txtMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Monto No es Inválido"
End If


strSQL = "select count(*) as Existe from Ins_Tipos_Seguros where Tipo_Seguro ='" & txtTipoSeguroCod.Text & "' and Activo = 1"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de seguro no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from Ins_Tipos_Cuentas where Tipo_Cuenta ='" & txtTipoCuentaCod.Text & "' and Activo = 1"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de Cuenta no se encuentra Activa!"
rs.Close


strSQL = "select count(*) as Existe from Ins_Vendedores where cod_vendedor = " & txtVendedorCod.Text & " and Activo = 1"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El vendedor no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from Socios where cedula = '" & txtCedula.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- La persona no existe en la base de datos!"
rs.Close


If Len(vMensaje) > 0 Then fxValida = False

End Function


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtPoliza.Text = "" Then txtPoliza.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 num_poliza from Ins_Polizas"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where num_poliza > '" & txtPoliza & "' order by num_poliza asc"
Else
   strSQL = strSQL & " where num_poliza < '" & txtPoliza & "' order by num_poliza desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  txtPoliza.Text = rs!Num_Poliza
  Call sbConsulta
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


Private Sub cmdAjustar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer

On Error GoTo vError

       
If Mid(txtEstado.Text, 1, 1) <> "A" Then
    MsgBox "No se puede modificar esta póliza porque no se encuentra Activa", vbExclamation
    Exit Sub
End If
       
strSQL = "update Ins_Polizas set cod_vendedor = " & txtVendedorCod.Text & ",Tipo_Seguro = '" & txtTipoSeguroCod.Text & "',Tipo_Cuenta = '" _
    & txtTipoCuentaCod.Text & "',notas = '" & txtNotas.Text & "',Monto = " & CCur(txtMonto.Text) & ", Cuota =  " & CCur(txtCuota.Text) _
    & ", Plazo = " & txtPlazo.Text & ", cedula = '" & Trim(txtCedula.Text) _
    & "' where num_poliza = '" & txtPoliza.Text & "'"
glogon.Conection.Execute strSQL

'Actualiza datos de Cobranza General y Variaciones en las Polizas
strSQL = "exec spInsCobrosActualiza"
glogon.Conection.Execute strSQL


MsgBox "Ajuste a Póliza realizado satisfactoriamente!", vbInformation
'TODO: Crear Bitácora

Call sbConsulta

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdEliminar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer

On Error GoTo vError

       
If Mid(txtEstado.Text, 1, 1) <> "A" Then
    MsgBox "No se puede Eliminar esta póliza porque no se encuentra Activa", vbExclamation
    Exit Sub
End If
       
strSQL = "exec spInsPolizaActivaElimina '" & txtPoliza.Text & "','" & glogon.Usuario & "'"
glogon.Conection.Execute strSQL
       
MsgBox "Póliza Eliminada / Ajustadas todas las referencias...", vbInformation
       
Call sbConsulta

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub FlatScrollBarTipoSeguro_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoSeguro.Tag = "" Then FlatScrollBarTipoSeguro.Tag = 0

strSQL = "select Top 1 Tipo_Seguro,Descripcion from Ins_Tipos_Seguros"

If FlatScrollBarTipoSeguro.Value > CLng(FlatScrollBarTipoSeguro.Tag) Then
   strSQL = strSQL & " where Tipo_Seguro > '" & txtTipoSeguroCod.Text & "' and Activo = 1 order by Tipo_Seguro asc"
Else
   strSQL = strSQL & " where Tipo_Seguro < '" & txtTipoSeguroCod.Text & "' and Activo = 1 order by Tipo_Seguro asc"
End If

FlatScrollBarTipoSeguro.Tag = FlatScrollBarTipoSeguro.Value

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
        txtTipoSeguroCod.Text = rs!Tipo_Seguro
        txtTipoSeguroDesc.Text = rs!Descripcion
Else
        txtTipoSeguroCod.Text = ""
        txtTipoSeguroDesc.Text = ""
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub FlatScrollBarTipoCuenta_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoCuenta.Tag = "" Then FlatScrollBarTipoCuenta.Tag = 0

strSQL = "select Top 1 Tipo_Cuenta,Descripcion from Ins_Tipos_Cuentas"

If FlatScrollBarTipoCuenta.Value > CLng(FlatScrollBarTipoCuenta.Tag) Then
   strSQL = strSQL & " where Tipo_Cuenta > '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by Tipo_Cuenta asc"
Else
   strSQL = strSQL & " where Tipo_Cuenta < '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by Tipo_Cuenta asc"
End If

FlatScrollBarTipoCuenta.Tag = FlatScrollBarTipoCuenta.Value

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  txtTipoCuentaCod.Text = rs!Tipo_Cuenta
  txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
 
 vModulo = 17

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpiaPantalla


End Sub

Private Sub sbLimpiaPantalla()
 

Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(11).Picture
ImgAutorizacion.ToolTipText = "Pendiente: Consulta/Nuevo"
 
 txtCedula = ""
 txtNombre = ""
 lblNombre.Caption = txtNombre.Text
 
 txtVendedorCod.Text = ""
 txtVendedorDesc.Text = ""

 txtTipoSeguroCod.Text = ""
 txtTipoSeguroDesc.Text = ""

 txtTipoCuentaCod.Text = ""
 txtTipoCuentaDesc.Text = ""
   
 txtEstado.Text = "Pendiente"
   
 txtMonto = "0"
 txtPlazo = "60"
 txtCuota = "0"
  

 txtNotas = ""
 

 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 StatusBarX.Panels(4).Text = ""
 StatusBarX.Panels(5).Text = ""

End Sub



Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

vPaso = True

strSQL = "select Pol.*,Ts.Descripcion as 'TipoSeguroDesc', Per.Nombre, isnull(Pol.Estado,'P') as 'Estado'" _
       & ",Ven.Nombre as 'VendedorNombre',Tc.descripcion as 'TipoCuentaDesc',dbo.MyGetdate() as 'FechaServer'" _
       & " from Ins_Polizas Pol inner join Ins_Tipos_Seguros Ts on Pol.Tipo_Seguro = Ts.Tipo_Seguro" _
       & " inner join Socios Per on Pol.cedula = Per.cedula" _
       & " left join Ins_Vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor" _
       & " left join Ins_Tipos_Cuentas Tc on Pol.Tipo_Cuenta = Tc.Tipo_Cuenta" _
       & " where Pol.num_poliza = '" & txtPoliza.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic, adLockOptimistic

If Not rs.EOF And Not rs.BOF Then
 
 vFecha = rs!FechaServer
 
 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 lblNombre.Caption = txtNombre.Text
 
 txtVendedorCod.Text = rs!cod_vendedor
 txtVendedorDesc.Text = rs!VendedorNombre
 
 txtTipoSeguroCod.Text = rs!Tipo_Seguro
 txtTipoSeguroDesc.Text = rs!TipoSeguroDesc
 txtTipoCuentaCod.Text = rs!Tipo_Cuenta
 txtTipoCuentaDesc.Text = rs!TipoCuentaDesc
 
 
 txtMonto.Text = Format(IIf(IsNull(rs!MONTO), 0, rs!MONTO), "Standard")
 txtPlazo.Text = rs!Plazo
 txtCuota.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 
 txtNotas = IIf(IsNull(rs!notas), "", rs!notas)



 Select Case rs!Estado
   Case "P" 'Pendiente
      txtEstado.Text = "Pendiente"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(7).Picture
      ImgAutorizacion.ToolTipText = "Activación: Pendiente"
      
   Case "A"
      txtEstado.Text = "Activa"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(5).Picture
      ImgAutorizacion.ToolTipText = "Póliza Activada!"
   Case "C"
      txtEstado.Text = "Cerrada"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(6).Picture
      ImgAutorizacion.ToolTipText = "Póliza Cerrada (Inactivada)"
  End Select

 StatusBarX.Panels(1).Text = rs!Registro_Usuario
 StatusBarX.Panels(2).Text = rs!registro_Fecha
 StatusBarX.Panels(3).Text = rs!Activa_Usuario & ""
 StatusBarX.Panels(4).Text = rs!Activa_Fecha & ""
 StatusBarX.Panels(5).Text = rs!Cierra_usuario & ""
 StatusBarX.Panels(6).Text = rs!Cierra_fecha & ""
 

Else
 If vEdita Then
    MsgBox "No existe la Póliza, verifique!", vbCritical
 End If
End If
rs.Close

vPaso = False

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub




Private Sub txtCuota_GotFocus()
On Error GoTo vError

txtCuota.Text = CCur(txtCuota.Text)

vError:
End Sub

Private Sub txtCuota_LostFocus()
On Error GoTo vError

txtCuota.Text = Format(CCur(txtCuota.Text), "Standard")

vError:

End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuota.SetFocus
End Sub

Private Sub txtPoliza_LostFocus()
  Call sbConsulta
End Sub

Private Sub txtTipoSeguroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Seguro"
   gBusquedas.Columna = "Tipo_Seguro"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Seguro,Descripcion  from Ins_Tipos_Seguros"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoSeguroCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from Ins_Tipos_Seguros where Tipo_Seguro = '" & txtTipoSeguroCod.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   txtTipoSeguroCod.Text = rs!Tipo_Seguro
   txtTipoSeguroDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub



Private Sub txtTipoSeguroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Seguro,Descripcion  from Ins_Tipos_Seguros"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If
End Sub


'--Tipo de Cuenta
Private Sub txtTipoCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Cuenta"
   gBusquedas.Columna = "Tipo_Cuenta"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Cuenta,Descripcion  from Ins_Tipos_Cuentas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoCuentaCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from Ins_Tipos_Cuentas where Tipo_Cuenta = '" & txtTipoCuentaCod.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   txtTipoCuentaCod.Text = rs!Tipo_Cuenta
   txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub



Private Sub txtTipoCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Cuenta,Descripcion  from Ins_Tipos_Cuentas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If
End Sub




Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub


Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:

End Sub



Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If


End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


txtNombre.Text = fxPersonaNombre(txtCedula)
lblNombre.Caption = txtNombre.Text


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If
End Sub

'--Vendedor
Private Sub txtVendedorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cod_Vendedor"
   gBusquedas.Columna = "Cod_Vendedor"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from Ins_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If


End Sub

Private Sub txtVendedorCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select Cod_Vendedor,Nombre from Ins_Vendedores where cod_Vendedor = " & txtVendedorCod.Text
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
    txtVendedorDesc.Text = rs!Nombre
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtVendedorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from Ins_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If
End Sub


Public Sub sbConsultaExterna(xOpTemp As String)
 txtPoliza.Text = xOpTemp
 Call sbConsulta
End Sub


Private Sub txtPoliza_Change()
 Call sbLimpiaPantalla

End Sub

Private Sub txtPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub




