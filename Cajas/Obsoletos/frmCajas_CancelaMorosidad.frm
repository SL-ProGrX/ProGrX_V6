VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajas_CancelaMorosidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cancelación Mora"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   12
      ToolTipText     =   "Código del Préstamo"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      ToolTipText     =   "Nombre Completo del Socio (Apellidos y Nombre)"
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox txtCedula 
      Height          =   315
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   10
      ToolTipText     =   "Campo para la Cédula de Identidad"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerarRecibo 
      Caption         =   "Aplicar"
      Enabled         =   0   'False
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
      Left            =   7680
      Picture         =   "frmCajas_CancelaMorosidad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Genera el recibo para validar el movimiento"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtId_solicitud 
      Height          =   315
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox cboTipoPago 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCajas_CancelaMorosidad.frx":0138
      Left            =   1080
      List            =   "frmCajas_CancelaMorosidad.frx":0148
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtAbono 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Presione ENTER para Activar el Boton de Aplicar"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.ComboBox cboTipo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCajas_CancelaMorosidad.frx":016F
      Left            =   1080
      List            =   "frmCajas_CancelaMorosidad.frx":0171
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ComboBox cboTipoAbono 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCajas_CancelaMorosidad.frx":0173
      Left            =   1080
      List            =   "frmCajas_CancelaMorosidad.frx":017D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer TimerVerificaPlanPagos 
      Left            =   8280
      Top             =   0
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Presione ENTER para Activar el Boton de Aplicar"
      Top             =   4800
      Width           =   1815
   End
   Begin MSComctlLib.ListView lswDetalle 
      Height          =   2535
      Left            =   0
      TabIndex        =   6
      Top             =   1845
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4471
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
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Operación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Fec.Proc"
         Object.Width           =   2011
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
         Text            =   "Int.Mor."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amortiza"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Cargo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblDesgloce 
      Height          =   330
      Left            =   6960
      TabIndex        =   25
      ToolTipText     =   "Detalle de pago"
      Top             =   5160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      ButtonWidth     =   1138
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Aplicar"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_CancelaMorosidad.frx":0193
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgBusqueda_Rapida 
      Height          =   255
      Index           =   2
      Left            =   8880
      Picture         =   "frmCajas_CancelaMorosidad.frx":02B7
      Stretch         =   -1  'True
      ToolTipText     =   "Busqueda Rápida"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblNombreCodigo 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2880
      TabIndex        =   24
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Image imgBusqueda_Rapida 
      Height          =   255
      Index           =   0
      Left            =   8880
      Picture         =   "frmCajas_CancelaMorosidad.frx":03B0
      Stretch         =   -1  'True
      ToolTipText     =   "Busqueda Rápida"
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Total Abono"
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
      Index           =   4
      Left            =   3360
      TabIndex        =   23
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Total Seleccionado"
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
      Index           =   0
      Left            =   3360
      TabIndex        =   22
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblId_solicitud 
      Caption         =   "lblId_solicitud"
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgBusqueda_Rapida 
      Height          =   240
      Index           =   1
      Left            =   5280
      Picture         =   "frmCajas_CancelaMorosidad.frx":04A9
      ToolTipText     =   "Busqueda Rápida"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblOpex 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   8160
      TabIndex        =   20
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle Cuotas Morosas"
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
      Height          =   290
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   9255
   End
   Begin VB.Label Label23 
      Caption         =   "Tipo - Pago"
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
      Left            =   0
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo - Doc"
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
      Left            =   0
      TabIndex        =   17
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo - Abono"
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
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operación"
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
      Index           =   0
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Línea"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cédula"
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
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmCajas_CancelaMorosidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMorosidad As New ADODB.Recordset
Dim mstrId_Moro As String
Dim mblnEdita As Boolean
Private miTran As Integer, vUltimoRecibo As Long
Dim mCurIntc As Currency, mCurIntm As Currency, mCurAmortiza As Currency, mCurCargo As Currency
Dim vRetencion

Private Sub Reporte1(strTitulo As String)

Dim Str As String, strRuta As String, strInicio As String, strFinal As String
Dim str1 As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Str = ""
str1 = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "CUOTAS MOROSAS"
 .ReportFileName = SIFGlobal.fxSIFPathReportes("crCM.rpt")
.Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
.Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
.Formulas(3) = "Titulo='" & strTitulo & "'"

Str = "{REG_CREDITOS.ID_SOLICITUD}=" & Trim(txtId_solicitud)
Str = Str & " and {REG_CREDITOS.ESTADO}='A'"
.SelectionFormula = Str


.SubreportToChange = "subMor"
Str = "{MOROSIDAD.ID_SOLICITUD} = {?Pm-REG_CREDITOS.ID_SOLICITUD} and {MOROSIDAD.ESTADO}='A'"

.SelectionFormula = Str

.PrintReport

End With
Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical

End Sub



Sub Limpia()
txtId_solicitud = ""
txtCedula = ""
txtNombre = ""
txtCodigo = ""
lblNombreCodigo = ""
lblId_Solicitud = ""
lswDetalle.ListItems.Clear
txtTotal.Text = ""
txtAbono.Text = ""
mstrId_Moro = ""
mblnEdita = True
txtAbono.Enabled = True
End Sub

Private Function fxAbonaMorosidad(curAbono As Currency, vID_Moro As Long, vFecha As Date) As Currency
'Devuelve el abono sobrante
Dim rs As New ADODB.Recordset, strSQL As String
Dim curRepartido(3) As Currency, vCuenta As String

curRepartido(0) = 0
curRepartido(1) = 0
curRepartido(2) = 0
curRepartido(3) = 0

strSQL = "Select * from morosidad where id_moro = " & vID_Moro
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then

 If curAbono >= rs!intm Then
    curRepartido(2) = rs!intm
    curAbono = curAbono - rs!intm
 Else
    curRepartido(2) = curAbono
    curAbono = 0
 End If

 If curAbono >= rs!intc Then
    curRepartido(1) = rs!intc
    curAbono = curAbono - rs!intc
 Else
    curRepartido(1) = curAbono
    curAbono = 0
 End If

 If curAbono >= rs!Cargo Then
    curRepartido(0) = rs!Cargo
    curAbono = curAbono - rs!Cargo
 Else
    curRepartido(0) = curAbono
    curAbono = 0
 End If



 If curAbono >= rs!Amortiza Then
    curRepartido(3) = rs!Amortiza
    curAbono = curAbono - rs!Amortiza
 Else
    curRepartido(3) = curAbono
    curAbono = 0
 End If

 'Suma varibles de modulo
 
 mCurCargo = mCurCargo + curRepartido(0)
 mCurIntc = mCurIntc + curRepartido(1)
 mCurIntm = mCurIntm + curRepartido(2)
 mCurAmortiza = mCurAmortiza + curRepartido(3)
 
 strSQL = "update morosidad set estado = 'C'," _
        & "fecult = '" & Format(vFecha, "yyyy/mm/dd") _
        & "',tcon = '20',ncon = '" & rs!ID_SOLICITUD _
        & "',abintc = " & curRepartido(1) _
        & ",abintm = " & curRepartido(2) _
        & ",abamortiza = " & curRepartido(3) _
        & ",abCargo = " & curRepartido(0) _
        & " where id_moro = " & rs!id_moro
 glogon.Conection.Execute strSQL
 
 If rs!intc + rs!intm + rs!Amortiza + rs!Cargo > curRepartido(1) _
    + curRepartido(2) + curRepartido(3) + curRepartido(0) Then
  'Insertar Registro con la diferencia
   strSQL = "insert morosidad(id_solicitud,codigo,fechap,intc,intm,amortiza," _
         & "estado,fecap,estadoi,fecult,cuota_morosa,cargo) values(" & rs!ID_SOLICITUD _
         & ",'" & rs!Codigo & "'," & rs!fechap & "," & rs!intc - curRepartido(1) _
         & "," & rs!intm - curRepartido(2) & "," & rs!Amortiza - curRepartido(3) _
         & ",'A'," & GLOBALES.glngFechaCR & ",'A','" & Format(vFecha, "yyyy/mm/dd") & "'," _
         & (rs!intc + rs!intm + rs!Amortiza + rs!Cargo) - (curRepartido(0) + curRepartido(1) + curRepartido(2) _
         + curRepartido(3)) & "," & rs!Cargo - curRepartido(0) & ")"
   glogon.Conection.Execute strSQL
 End If

End If
rs.Close
fxAbonaMorosidad = curAbono

End Function

Private Sub sbCargaGridMorosidad()
Dim curAux As Currency, Str As String
Dim strSQL As String, i As Integer
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from MOROSIDAD where estado = 'A'" _
       & " and id_solicitud=" & lblId_Solicitud & " order by fechap asc"

lswDetalle.ListItems.Clear
rsMorosidad.CursorLocation = adUseServer
rsMorosidad.Open strSQL, glogon.Conection, adOpenStatic

On Error Resume Next

With rsMorosidad
   Do While Not .EOF
     Set itmX = lswDetalle.ListItems.Add(lswDetalle.ListItems.Count + 1, , rsMorosidad!id_moro)
         itmX.Tag = itmX.Index
         itmX.SubItems(1) = rsMorosidad!ID_SOLICITUD
         itmX.SubItems(2) = Format(rsMorosidad!fechap, "####-##")
         itmX.SubItems(3) = Format(rsMorosidad!intc, "Standard")
         itmX.SubItems(4) = Format(rsMorosidad!intm, "Standard")
         itmX.SubItems(5) = Format(rsMorosidad!Amortiza, "Standard")
         itmX.SubItems(6) = Format(rsMorosidad!Cargo, "Standard")
         itmX.SubItems(7) = Format(rsMorosidad!intc + rsMorosidad!intm + rsMorosidad!Amortiza + rsMorosidad!Cargo, "Standard")
    .MoveNext
   Loop
End With

rsMorosidad.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
End Sub

Function ExisteNombre(strCedula As String)

Dim strSQL As String
Dim recSocios As New ADODB.Recordset
ExisteNombre = False

On Error GoTo vError

strSQL = "Select nombre from SOCIOS where cedula='" & strCedula & "'"
With recSocios
.Open strSQL, glogon.Conection, adOpenStatic
If Not .EOF And .RecordCount >= 1 Then
  ExisteNombre = True
  txtNombre = !Nombre
End If
.Close
End With

Exit Function
vError:
MsgBox Err.Description, vbCritical
End Function

Function fxExisteRegistro(strCodigo As String, i As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxExisteRegistro = ""

If i = 2 Then
    txtId_solicitud = ""
    lswDetalle.ListItems.Clear
    txtTotal.Text = ""
    txtAbono.Text = ""
    mstrId_Moro = ""
    mblnEdita = True
    txtAbono.Enabled = True
    cmdGenerarRecibo = False
    txtId_solicitud.Enabled = False
    txtCedula.Enabled = False
    txtCodigo.Enabled = False
    imgBusqueda_Rapida.Item(0).Enabled = False
    imgBusqueda_Rapida.Item(1).Enabled = False
    imgBusqueda_Rapida.Item(2).Enabled = False
ElseIf i = 1 Then
    txtCedula = ""
    txtCodigo = ""
    txtNombre = ""
    txtProceso.Text = ""
    txtProceso.Tag = ""
    lblNombreCodigo = ""
    lswDetalle.ListItems.Clear
    txtTotal.Text = ""
    txtAbono.Text = ""
    mstrId_Moro = ""
    mblnEdita = True
    txtAbono.Enabled = True
    cmdGenerarRecibo = False
    txtId_solicitud.Enabled = False
    txtCedula.Enabled = False
    txtCodigo.Enabled = False
    imgBusqueda_Rapida.Item(0).Enabled = False
    imgBusqueda_Rapida.Item(1).Enabled = False
    imgBusqueda_Rapida.Item(2).Enabled = False
End If

strSQL = "Select R.*,C.descripcion,C.retencion,C.poliza,S.nombre" _
       & " from REG_CREDITOS R inner join CATALOGO C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula where"

If Trim(strCodigo) <> "" And Trim(txtCedula) <> "" Then
    strSQL = strSQL & " R.cedula='" & Trim(txtCedula) & "' and R.Codigo='" & strCodigo & "'"
ElseIf Trim(txtId_solicitud) <> "" Then
    strSQL = strSQL & " R.Id_solicitud=" & Trim(txtId_solicitud)
End If

strSQL = strSQL & " and R.estado='A'"

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And rs.RecordCount >= 1 Then
  lblNombreCodigo = rs!Descripcion
  txtId_solicitud = rs!ID_SOLICITUD
  txtCodigo = rs!Codigo
  txtCedula = rs!Cedula
  txtNombre = rs!Nombre
  
    txtProceso.Tag = rs!Proceso
    Select Case rs!Proceso
      Case "N"
        txtProceso.Text = "Normal"
      Case "T"
        txtProceso.Text = "Traspaso Deuda"
      Case "J"
        txtProceso.Text = "Cobro Judicial"
      Case "I"
        txtProceso.Text = "Incobrable"
    End Select
    
  lblOpex.Caption = IIf(rs!opex = 1, "OPEX", "")
  fxExisteRegistro = rs!ID_SOLICITUD
  
  If rs!retencion = "S" Or rs!Poliza = "S" Then
    vRetencion = True
  Else
    vRetencion = False
  End If
  
End If
rs.Close

Exit Function

vError:
    MsgBox Err.Description, vbCritical

End Function



Private Sub cboTipoAbono_Click()
Dim vValor As Boolean, i As Integer
If cboTipoAbono.Text = "Seleccion" Then
  vValor = False
Else
  vValor = True
End If

On Error Resume Next

For i = 1 To lswDetalle.ListItems.Count
  lswDetalle.SelectedItem = lswDetalle.ListItems.Item(i)
  lswDetalle.SelectedItem.Checked = vValor
Next i

Call lswDetalle_Click


End Sub

Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 cmdGenerarRecibo.Enabled = True
 cmdGenerarRecibo.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()

txtId_solicitud.Enabled = True
txtCedula.Enabled = True
txtCodigo.Enabled = True
cmdGenerarRecibo.Enabled = False
imgBusqueda_Rapida.Item(0).Enabled = True
imgBusqueda_Rapida.Item(1).Enabled = True
imgBusqueda_Rapida.Item(2).Enabled = True
Call Limpia
Call RefrescaTags(Me)
End Sub

Private Sub EstableceUltimoAbono()
Dim strSQL As String, rs As New ADODB.Recordset

'Ultimo Abono Registrado.
strSQL = "select max(fechap) as Corte from morosidad where estado = 'C' and id_solicitud = " & txtId_solicitud
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   strSQL = "update reg_creditos set fecult = " & rs!Corte _
          & " where id_solicitud = " & txtId_solicitud & " and fecult < " & rs!Corte
   glogon.Conection.Execute strSQL
End If
rs.Close

End Sub

 
Private Sub cmdGenerarRecibo_Click()
Dim curAbono As Currency, vPaso As Boolean, i As Integer
Dim rs As New ADODB.Recordset, strSQL As String, iCuotas As Integer
Dim vFecha  As Date, lngRecibo As Long, vCuenta As String, vTipo As String
Dim vConcepto As String


'Verifica el proceso
If txtProceso.Tag = "J" Then
   If Not fxCRDAbonosAutorizados(txtCodigo.Text, txtProceso.Tag) Then
      MsgBox "- El usuario actual no cuenta con permisos para realizar abonos a Creditos en Cobro Judicial, verifique...", vbExclamation
      Exit Sub
   End If
End If

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
  MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
  Exit Sub
End If


Me.MousePointer = vbHourglass


If Not IsNumeric(txtAbono.Text) Then
 MsgBox "Verifique el Valor del Abono..." & txtAbono.Text, vbOKOnly
 Me.MousePointer = vbDefault
 Exit Sub
End If
 
If CCur(txtAbono) > CCur(txtTotal.Text) Then
 MsgBox "El abono es mayor a las cuotas... " & txtAbono.Text, vbOKOnly
 Me.MousePointer = vbDefault
 Exit Sub
End If
 
vUltimoRecibo = 0
vConcepto = "CRD001"

If Trim(txtAbono.Text) <> "" And Trim(txtTotal.Text) <> "" Then
        
    lngRecibo = 0
    vFecha = fxFechaServidor
    
    'Configuracion del Documento
    'vTipo = fxTipoASEDoc(cboTipo.Text)
    vTipo = SIFGlobal.fxSIFCodText(cboTipo)
    vCuenta = Trim(fxDocumentoCuenta(vTipo))
    lngRecibo = fxDocumentoConsecutivo(vTipo)
    
    vUltimoRecibo = lngRecibo
    
    If vAseDocValido = False Then
        Me.MousePointer = vbDefault
        MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
              & " válida para esta operación...", vbCritical
        Exit Sub
    End If
    
    
    curAbono = CCur(txtAbono.Text)
    vPaso = False
    mCurAmortiza = 0
    mCurIntc = 0
    mCurIntm = 0
    mCurCargo = 0
    iCuotas = 0
    
''Eliminar
'    curAbono = 17400.73
'    vPaso = False
'    mCurAmortiza = 8954.56
'    mCurIntc = 8111.62
'    mCurIntm = 334.55
'    mCurCargo = 0
'    iCuotas = 2

On Error GoTo vError

'Inicia Transacciones
glogon.Conection.BeginTrans


    For i = 1 To lswDetalle.ListItems.Count
      lswDetalle.SelectedItem = lswDetalle.ListItems.Item(i)
      If lswDetalle.SelectedItem.Checked = True Then
       If curAbono > 0 Then
        iCuotas = iCuotas + 1
        curAbono = fxAbonaMorosidad(curAbono, lswDetalle.SelectedItem.Text, vFecha)
        vPaso = True
       End If
      End If
    Next i
   
    If uRecibos Then lngRecibo = fxDocumentoAbono(vTipo, CStr(lngRecibo), vConcepto, vCuenta, mCurIntc, mCurIntm, mCurAmortiza, mCurCargo, iCuotas)
    
    'actualizar reg_creditos
    If Not vRetencion Then
        strSQL = "update reg_creditos set saldo = saldo - " & mCurAmortiza _
               & ",amortiza = amortiza + " & mCurAmortiza & ",saldo_mes = saldo_mes - " _
               & mCurAmortiza & ",interesc = interesc + " & mCurIntc + mCurIntm _
               & " where id_solicitud = " & txtId_solicitud.Text
    Else 'Retencion
        strSQL = "update reg_creditos set amortiza = amortiza + " & mCurAmortiza _
               & ",interesc = interesc + " & mCurIntc + mCurIntm _
               & " where id_solicitud = " & txtId_solicitud.Text
    End If
    
    glogon.Conection.Execute strSQL
    
    strSQL = "update morosidad set tcon = '" & fxTipoASENumero(SIFGlobal.fxSIFCodText(cboTipo.Text)) _
           & "',ncon = '" & IIf((lngRecibo = 0), "null", lngRecibo) & "', Usuario = '" & glogon.Usuario & "', Cod_Concepto = '" & vConcepto & "', cod_caja = ''" _
           & " where tcon = '20' and ncon = '" & txtId_solicitud.Text & "'"
    glogon.Conection.Execute strSQL
   

'Cierra Transacciones
glogon.Conection.CommitTrans

Call Bitacora("Modifica", "Abona Morosidad Op: " & txtId_solicitud _
        & " Cuota:" & iCuotas & " amortiza: " & mCurAmortiza & " int: " _
        & mCurIntc + mCurIntm)


Me.MousePointer = vbDefault
MsgBox "Abono Realizado... " & cboTipo.Text & " #" & vUltimoRecibo, vbInformation


Call EstableceUltimoAbono


  If vUltimoRecibo > 0 And vPaso Then _
    Call sbImprimeRecibo(vUltimoRecibo, fxTipoASEDoc(SIFGlobal.fxSIFCodText(cboTipo.Text)))
    
    cmdGenerarRecibo.Enabled = False
    txtAbono.Enabled = False
    mblnEdita = False

End If

txtAbono = ""
txtTotal = ""

Call sbCargaGridMorosidad
Call RefrescaTags(Me)
cmdGenerarRecibo.Enabled = False
Exit Sub

vError:
    glogon.Conection.RollbackTrans
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical

    
End Sub


Private Sub cmdReporte_Click()

If Trim(txtId_solicitud) <> "" Then
Call Reporte1("CUOTAS MOROSAS")
End If

End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub


Private Sub Form_Load()
Dim strSQL As String
 vModulo = 5

 If GLOBALES.SysPlanPagos = 1 Then
    TimerVerificaPlanPagos.Interval = 10
 Else
   'Carga Load Normalmente
   Call Formularios(Me)
   'strSQL = "select tipo_documento + ' - ' + Descripcion as itmx from sif_documentos where activo = 1"
   
   strSQL = "select C.tipo_documento + ' - ' + D.Descripcion as itmx  from SIF_DOCUMENTOS D" _
        & " inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
        & " AND  C.cod_caja =  '" & ModuloCajas.mCaja & "' order by C.tipo_documento"
        
   Call sbLlenaCbo(cboTipo, strSQL, False, False)
'    cboTipoPago.Text = "Efectivo"
'    cboTipoAbono.Text = "Seleccion"
    
'     cboTipo.AddItem "Recibo"
'     cboTipo.AddItem "Nota Credito"
'     cboTipo.Text = "Recibo"
     
    mblnEdita = True
    
    Call Formularios(Me)
    Call RefrescaTags(Me)
    cmdGenerarRecibo.Enabled = False
 End If

End Sub
Private Sub imgBusqueda_Rapida_Click(Index As Integer)
Dim strId_solicitud As String

Dim i As Integer

On Error GoTo vError

Select Case Index

  Case 0

        Call Limpia
        
        gBusquedas.Convertir = "N"
        
        gBusquedas.Consulta = "Select Reg_creditos.Cedula,id_solicitud,Nombre,Codigo from REG_CREDITOS,SOCIOS"
        gBusquedas.Columna = "Reg_Creditos.CEDULA"
        gBusquedas.Orden = "Reg_Creditos.CEDULA"
        gBusquedas.Filtro = " AND REG_CREDITOS.ESTADO = 'A' AND REG_CREDITOS.PROCESO <> 'J' AND REG_CREDITOS.cedula=SOCIOS.cedula"
        
        frmBusquedas.Show vbModal
        
        txtCedula = Trim(gBusquedas.Resultado)
        gBusquedas.Consulta = ""
        gBusquedas.Columna = ""
        gBusquedas.Orden = ""
        gBusquedas.Resultado = ""
        gBusquedas.Filtro = ""
        
        If Trim(txtCedula) <> "" Then
            If ExisteNombre(txtCedula) Then
            End If
        End If

  Case 1

        gBusquedas.Convertir = "S"
        gBusquedas.Consulta = "Select id_solicitud,codigo,Reg_creditos.Cedula,Nombre from REG_CREDITOS,SOCIOS"
        gBusquedas.Columna = "Id_Solicitud"
        gBusquedas.Orden = "Id_Solicitud"
        gBusquedas.Filtro = " AND REG_CREDITOS.ESTADO = 'A' REG_CREDITOS.PROCESO <> 'J' AND REG_CREDITOS.cedula=SOCIOS.cedula"
        
        frmBusquedas.Show vbModal
        
        txtId_solicitud = Trim(gBusquedas.Resultado)
        gBusquedas.Consulta = ""
        gBusquedas.Columna = ""
        gBusquedas.Orden = ""
        gBusquedas.Resultado = ""
        gBusquedas.Filtro = ""
        
        If Trim(txtId_solicitud) <> "" Then
         strId_solicitud = fxExisteRegistro(txtCodigo, 1)
         If strId_solicitud <> "" Then
               lblId_Solicitud = strId_solicitud
         End If
        End If


  Case 2

        gBusquedas.Convertir = "N"
        gBusquedas.Consulta = "Select distinct REG_CREDITOS.CODIGO,CATALOGO.DESCRIPCION from REG_CREDITOS,CATALOGO"
        gBusquedas.Columna = "REG_CREDITOS.CODIGO"
        gBusquedas.Orden = "REG_CREDITOS.CODIGO"
        gBusquedas.Filtro = " AND REG_CREDITOS.CODIGO=CATALOGO.CODIGO"
        
        frmBusquedas.Show vbModal
        
        txtCodigo = Trim(gBusquedas.Resultado)
        gBusquedas.Consulta = ""
        gBusquedas.Columna = ""
        gBusquedas.Orden = ""
        gBusquedas.Resultado = ""
        
        lblId_Solicitud = ""
        txtId_solicitud = ""
        
        If Trim(txtCodigo) <> "" And Trim(txtCedula) <> "" Then
         strId_solicitud = fxExisteRegistro(txtCodigo, 2)
         If strId_solicitud <> "" Then
               lblId_Solicitud = strId_solicitud
         End If
        End If

End Select

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub



Private Sub lswDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 Set Conlsw.lswX = lswDetalle
 Conlsw.Abre
End If
End Sub
Private Sub lswDetalle_Click()
Dim curTotalPagar As Currency, i As Integer

On Error Resume Next
curTotalPagar = 0
For i = 1 To lswDetalle.ListItems.Count
  lswDetalle.SelectedItem = lswDetalle.ListItems.Item(i)
  If lswDetalle.SelectedItem.Checked = True Then
     curTotalPagar = curTotalPagar + CCur(lswDetalle.SelectedItem.SubItems(3)) _
     + CCur(lswDetalle.SelectedItem.SubItems(4)) _
     + CCur(lswDetalle.SelectedItem.SubItems(5)) _
     + CCur(lswDetalle.SelectedItem.SubItems(6))
  End If
Next i

txtTotal.Text = Format(curTotalPagar, "Standard")
txtAbono.SetFocus

End Sub

Private Sub tblDesgloce_ButtonClick(ByVal Button As MSComctlLib.Button)
ModuloCajas.mTotalAplicar = txtTotal
ModuloCajas.mCliente = txtCedula.Text & "-" & txtNombre.Text
ModuloCajas.mServicio = "Cancelación en Mora"
frmCajas_DetallePago.Show vbModal

    
If ModuloCajas.mTiquete = Empty Then
    cmdGenerarRecibo.Enabled = False
    MsgBox "Aún no ha desglosado la transacción"
    Exit Sub
Else
    cmdGenerarRecibo.Enabled = True
    Call RefrescaTags(Me)
    txtAbono.Text = Format(ModuloCajas.mTotalAplicar, "Standard")
End If
End Sub

Private Sub TimerVerificaPlanPagos_Timer()

TimerVerificaPlanPagos.Interval = 0
Call sbSIFForms("frmCR_AbonosNew", 0, , , False)
Unload Me

End Sub



Private Sub txtAbono_GotFocus()
On Error Resume Next
txtAbono = CCur(txtAbono)
End Sub

Private Sub txtAbono_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtAbono = Format(txtAbono, "###,###,###,##0.00")
  cboTipoPago.SetFocus
End If
End Sub

Private Sub txtAbono_LostFocus()
'If Trim(txtFechaCuota) <> "" And Trim(txtTotal.Text) <> "" And Trim(txtAbono.Text) <> "" And mblnEdita = True Then
 cmdGenerarRecibo.Enabled = True
'End If
Call RefrescaTags(Me)
End Sub
Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If Trim(txtCedula) <> "" And KeyAscii = vbKeyReturn Then
    If ExisteNombre(txtCedula) Then
      txtCodigo.SetFocus
    End If
End If
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Dim strId_solicitud As String


On Error GoTo vError
Me.MousePointer = vbHourglass
If Trim(txtCodigo) <> "" And Trim(txtCedula) <> "" And KeyAscii = vbKeyReturn Then
 strId_solicitud = fxExisteRegistro(txtCodigo, 2)
 If strId_solicitud <> "" Then
       lblId_Solicitud = strId_solicitud
 End If
End If
Me.MousePointer = vbDefault

vError:
Me.MousePointer = vbDefault


End Sub

Private Sub txtFechaCuota_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) Then
   KeyAscii = 0
 End If
End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtId_solicitud = xOpTemp
 Call txtId_solicitud_KeyPress(vbKeyReturn)
End Sub


Private Sub txtId_solicitud_KeyPress(KeyAscii As Integer)
Dim strId_solicitud As String

On Error GoTo vError


Me.MousePointer = vbHourglass

If Trim(txtId_solicitud) <> "" And KeyAscii = vbKeyReturn Then
    strId_solicitud = fxExisteRegistro(txtCodigo, 1)
    If strId_solicitud <> "" Then
       lblId_Solicitud = strId_solicitud
       Call sbCargaGridMorosidad
    End If
End If

vError:
Me.MousePointer = vbDefault

End Sub

Private Function fxDocumentoAbono(pTipoDoc As String, pComprobante As String, pConcepto As String, pCuenta As String _
                                , curIntC As Currency, curIntM As Currency, curAmortiza As Currency, curCargo As Currency _
                                , iCuotas As Integer) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim rs2 As New ADODB.Recordset, strCliente As String, lngRecibo As Long

lngRecibo = CLng(pComprobante)
fxDocumentoAbono = lngRecibo
 
 
 'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtId_solicitud.Text
rs.Open strSQL, glogon.Conection, adOpenStatic


rs2.CursorLocation = adUseServer
rs2.Open "select saldo from reg_creditos where id_solicitud = " & txtId_solicitud, glogon.Conection, adOpenStatic
    strLinea(1) = "Saldo Anterior    " & Format(rs2!Saldo, "Standard")
    strLinea(2) = "Interes Corriente " & Format(curIntC, "Standard")
    strLinea(3) = "Interes Moratorio " & Format(curIntM, "Standard")
    strLinea(4) = "Amortizacion      " & Format(curAmortiza, "Standard")
    If vRetencion Then
        strLinea(5) = "Saldo Actual      " & Format(rs2!Saldo, "Standard")
    Else
        strLinea(5) = "Saldo Actual      " & Format(rs2!Saldo - curAmortiza, "Standard")
    End If
    strLinea(6) = "Operación         " & txtId_solicitud
    strLinea(7) = "Línea             " & txtCodigo & "-" & UCase(lblOpex.Caption)
    strLinea(8) = "Cargo             " & Format(curCargo, "Standard")
    strLinea(9) = "Usuario           " & glogon.Usuario
    strLinea(10) = "Cuotas Abonadas   " & iCuotas
rs2.Close

If GLOBALES.SysDocVersion = 1 Then

        strCliente = Trim(txtCedula) & " - " & Trim(txtNombre)
        strCliente = Mid(strCliente, 1, 45)
        
        
        strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,dp)" _
                & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") & "','" & strCliente & "','" _
                & "ABONO A CUOTA EN MORA Op:" & txtId_solicitud & "'," & curIntC + curIntM + curAmortiza & ",'" & glogon.Usuario & "','P','" _
                & fxTipoPago(cboTipoPago.Text) & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                & vAseDocDetalle & "','" & vAseDocDeposito & "')"
        
        glogon.Conection.Execute strSQL
        
        'ASIENTO
        If curCargo > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & fxCBRParametro("23") & "'," & curCargo & ",'H')"
          glogon.Conection.Execute strSQL
        End If
        
        If curIntC > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Trim(rs!ctaintc) & "'," & curIntC & ",'H')"
          glogon.Conection.Execute strSQL
        End If
        
        If curIntM > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Trim(rs!ctaintc) & "'," & curIntM & ",'H')"
          glogon.Conection.Execute strSQL
        End If
        
        If curAmortiza > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & curAmortiza & ",'H')"
          glogon.Conection.Execute strSQL
        End If
        
        If curIntC + curIntM + curAmortiza + curCargo > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & pCuenta & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'D')"
          glogon.Conection.Execute strSQL
        End If
        
Else
  'Control de Documentos v2
   
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
                & " values('" & lngRecibo & "','" & pTipoDoc & "',getdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
                & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtId_solicitud.Text _
                & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                & vAseDocDetalle & "','" & vAseDocDeposito & "')"
        glogon.Conection.Execute strSQL
        
        'ASIENTO
        If curIntC + curIntM + curAmortiza + curCargo > 0 Then
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC + curIntM + curCargo + curAmortiza & ",'D','" & rs!cod_divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & pCuenta _
                 & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          glogon.Conection.Execute strSQL
          
          ''proceso pra crear asiento forma pago

          strSQL = "exec dbo.spCajasPorcesos '" & ModuloCajas.mTiquete & "','" & pTipoDoc & "','" & pConcepto & "'" _
             & ",'" & lngRecibo & "'," & CCur(curIntC + curIntM + curAmortiza + curCargo) & "" _
             & " ,'" & ModuloCajas.mOficina & "','" & GLOBALES.gOficinaCentroCosto & "'," & GLOBALES.gEnlace & "" _
             & "," & ModuloCajas.mApertura & ",'" & ModuloCajas.mCaja & "','" & Trim(txtCedula.Text) & "'"
             
          glogon.Conection.Execute strSQL
          'en caso que se haya utilizado saldos a favor
          If ModuloCajas.mCasosSFAplicados > 0 Then Call sbActualizaSaldosFavor(ModuloCajas.mCasosSFAplicados, txtCedula.Text)
          
        End If
        
        
'        If curIntC > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If
'
'        If curIntM > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntM & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If
'
'        If curCargo > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curCargo & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!CtaCargos _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If
'
'
'        If curAmortiza > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If

      
      

End If
rs.Close

End Function




