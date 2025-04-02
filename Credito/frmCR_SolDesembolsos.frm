VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmCR_SolDesembolsos 
   Caption         =   "Desembolsos"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frmCR_SolDesembolsos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B.SSDBGrid ssGriddgd 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   6975
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   0
      HeadFont3D      =   1
      Font3D          =   1
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   12303
      _ExtentY        =   4260
      _StockProps     =   79
      Caption         =   "Desembolsos de la solicitud"
   End
   Begin VB.TextBox txtmonto 
      Height          =   285
      Left            =   1200
      MaxLength       =   11
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtConcepto 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1200
      MaxLength       =   35
      TabIndex        =   2
      ToolTipText     =   "concepto del desembolso"
      Top             =   720
      Width           =   5535
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   1005
      ButtonWidth     =   1561
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Insertar"
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Borrar"
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Deshacer"
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ayuda"
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cerrar"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cierra esta ventana"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   7440
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":0BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":14C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":17DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":23D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":26F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":2A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SolDesembolsos.frx":32EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medCuenta 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Caption         =   "Disponible"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblMontoDisponible 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image imgBusqueda_Rapida 
      Height          =   255
      Index           =   0
      Left            =   6840
      Picture         =   "frmCR_SolDesembolsos.frx":3606
      Stretch         =   -1  'True
      ToolTipText     =   "Busqueda Rápida"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblCuenta 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Cuenta Conta."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Labdesembolsos 
      Caption         =   "Labdesembolsos"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Monto"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmCR_SolDesembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnmodificar As Boolean
Dim mboton As MSComctlLib.Button
Dim rsCuentas As New ADODB.Recordset
Dim intCaracteres As Integer 'Almacena el numero total de caracteres de la mascara
Dim mdblTotalDesembolsos As Double
Dim mcurMontoApr As Currency
Private curTotalDesRefu As Currency
Private mcurPrimerAbono As Currency
Private mcurIntHF As Currency
Private mblnFormato As Currency

Private Sub RefrescaGrid()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select * from desembolsos where id_solicitud = " & Operacion.Operacion
With rs
  .Open strSQL, GLOBALES.gConDatos, adOpenStatic
  If Not .EOF And .RecordCount >= 1 Then
   Call CargaGrid(rs, ssGriddgd)
  Else
   Call CargaGrid(rs, ssGriddgd)
  End If
  .Close
End With

End Sub
Private Sub CargaGrid(rs As Recordset, ByRef ssGrid As SSDBGrid)
Dim str As String
Dim strSQL As String
Dim i As Integer

ssGrid.RemoveAll
On Error Resume Next
ssGrid.Redraw = False
rs.MoveFirst
With rs
   Do While .EOF = False
    str = ""
    For i = 0 To .Fields.Count - 1
     If i = 4 Then
         str = str & Format(.Fields(i).Value, "###,###,##0.00") & vbTab
     Else
          str = str & .Fields(i).Value & vbTab
     End If
    Next i
    
    ssGrid.AddItem str
    .MoveNext
   Loop
End With
ssGrid.Redraw = True
rs.Close 'quito comentario 01/07/1999

End Sub
Function Valida()
Valida = True
If Trim(txtConcepto) = "" Or Trim(txtmonto) = "" Or Trim(medCuenta.Text) = "" Then
 MsgBox "Faltan datos", vbOKOnly
 Valida = False
End If
End Function
Sub Inactu()

Dim icomite As Integer
Dim iConsec As Long
Dim strSQL As String, rec As New ADODB.Recordset, str As String
Dim iTrans As Integer
On Error GoTo CapturaError

If mblnmodificar = False And Trim(Labdesembolsos.Caption) <> "Labdesembolsos" And Trim(Labdesembolsos.Caption) <> "" Then
    strSQL = "Select * from desembolosos Where id_desembolso =" & Trim(Labdesembolsos.Caption)
   With rec
    .ActiveConnection = GLOBALES.gConDatos
    .CursorType = adOpenStatic
    .Source = strSQL
    .Open
    If Not .EOF And .RecordCount >= 1 Then
      Me.MousePointer = vbDefault
      MsgBox "Codigo ya existe: " & !id_desembolso, vbOKOnly
      Exit Sub
    End If
    .Close
   End With
End If

If mblnmodificar = False Then
   iConsec = 0
   strSQL = "Select Max(id_desembolso) as iconsec from desembolsos "
   With rec
    .ActiveConnection = GLOBALES.gConDatos
    .CursorType = adOpenStatic
    .Source = strSQL
    .Open
    If Not .EOF And Not IsNull(!iConsec) Then
      iConsec = !iConsec
    End If
    .Close
   End With
   Labdesembolsos.Caption = iConsec + 1
   
str = "insert into desembolsos(codigo,id_solicitud,concepto,monto,"
str = str & "cuenta_conta)values('" & Trim(frmCR_SolicitudesFormalizacion.txtCodigoCredito.Text) & "',"
str = str & Trim(frmCR_SolicitudesFormalizacion.txtNumeroSolicitud.Text) & ","
str = str & "'" & UCase(Trim(txtConcepto)) & "',"
str = str & Format(txtmonto, "########0.00") & ","
str = str & "'" & Trim(medCuenta.Text) & "')"
GLOBALES.gConDatos.Execute str

Else  'si modificar una existente

str = "update desembolsos set codigo='"
str = str & Trim(frmCR_SolicitudesFormalizacion.txtCodigoCredito.Text) & "',"
str = str & "id_solicitud=" & Trim(frmCR_SolicitudesFormalizacion.txtNumeroSolicitud.Text) & ","
str = str & "concepto='" & UCase(Trim(txtConcepto)) & "',"
str = str & "monto=" & Format(txtmonto, "########0.00") & ","
str = str & "cuenta_conta='" & Trim(medCuenta.Text) & "'"
str = str & " where id_desembolso=" & Trim(Labdesembolsos.Caption)
GLOBALES.gConDatos.Execute str
 
End If

Exit Sub
CapturaError:
Me.MousePointer = vbDefault
Call ProcedimientoErrores(Me.Name)


End Sub
Sub habentrada()
 txtConcepto.Enabled = True
 txtmonto.Enabled = True
 medCuenta.Enabled = True
End Sub
Sub DeshabEntrada()
 txtConcepto.Enabled = False
 txtmonto.Enabled = False
 medCuenta.Enabled = False
End Sub
Sub Limpia()
 txtConcepto.Text = ""
 txtmonto.Text = ""
 medCuenta.Text = ""
 lblCuenta.Caption = ""
 Labdesembolsos.Caption = ""

'buscar la cuenta

End Sub
Sub barra1(tlb As Toolbar, opcion As Integer, ByVal Button As MSComctlLib.Button)

With tlb.Buttons
Select Case opcion
    Case 2 'Inicializa la barra
               Call DeshabEntrada
               .Item(1).Enabled = True  'Insertar
               .Item(2).Enabled = True  'Modificar
               .Item(3).Enabled = True  'Borrar
               .Item(4).Enabled = False 'Salvar
               .Item(5).Enabled = False 'Deshacer
               .Item(6).Enabled = True  'Ayuda
               .Item(7).Enabled = True  'Salir
End Select 'opcion
End With

End Sub
Function Consulta(strCodcre As String, strCodSol As String)
Dim i As Integer, strCodigo As String, strGarantia As String, strTramite As String
Dim strSQL As String, rec As New ADODB.Recordset
Consulta = False

On Error GoTo CapturaError

If Trim(strCodSol) <> "" Then
    strSQL = "Select * from desembolsos where codigo='" & Trim(strCodcre) & "'"
    strSQL = strSQL & " and id_solicitud=" & "" & strCodSol
    With rec
    .ActiveConnection = GLOBALES.gConDatos.ConnectionString
    .CursorType = adOpenStatic
    .Source = strSQL
    .Open
    If Not .EOF And .RecordCount >= 1 Then
     Consulta = True
     Call CargaGrid(rec, ssGriddgd)
    Else
     Consulta = False
    End If
    End With
Else
    MsgBox "Digite código de Solicitud"
End If

Exit Function
CapturaError:
Call ProcedimientoErrores(Me.Name)


End Function
Private Sub Form_Load()
Dim mbln As Boolean

GLOBALES.gintModulo = 3
GLOBALES.gstrFormCargado = Me.Name
Call Formularios

Call Iconos_ToolBar(tlbPrincipal)

Call CargaCuenta

rsCuentas.Source = "select * from empresas"
rsCuentas.ActiveConnection = GLOBALES.gConDatos
rsCuentas.CursorType = adOpenStatic
rsCuentas.Open

With rsCuentas
 intCaracteres = !nivel1
 intCaracteres = intCaracteres + !nivel2
 intCaracteres = intCaracteres + !nivel3
 intCaracteres = intCaracteres + !nivel4
 intCaracteres = intCaracteres + !nivel5
 .Close
End With

With ssGriddgd
    .Columns.Add 0
    .Columns(0).Caption = "ID"
    .Columns(0).Width = 1200
    .Columns.Add 1
    .Columns(1).Caption = "# Operación"
    .Columns(1).Width = 1200
    .Columns.Add 2
    .Columns(2).Caption = "Código"
    .Columns(2).Width = 900
    .Columns.Add 3
    .Columns(3).Caption = "Concepto"
    .Columns(3).Width = 3200
    .Columns.Add 4
    .Columns(4).Caption = "Monto"
    .Columns(4).Width = 1200
    .Columns.Add 5
    .Columns(5).Caption = "Cuenta"
    .Columns(5).Width = 900
End With

Call frmCR_SolicitudesFormalizacion.TotalDesRefu(curTotalDesRefu)
mcurMontoApr = frmCR_SolicitudesFormalizacion.txtMontoAprobado
mcurPrimerAbono = frmCR_SolicitudesFormalizacion.PrimerAbono
mcurIntHF = Format(frmCR_SolicitudesFormalizacion.IntHastaFormalizar, "########0.00")
lblMontoDisponible = Format(mcurMontoApr - curTotalDesRefu - mcurPrimerAbono - mcurIntHF, "standard")

If Trim(frmCR_SolicitudesFormalizacion.txtCodigoCredito) <> "" And Trim(frmCR_SolicitudesFormalizacion.txtNumeroSolicitud) <> "" Then
 mbln = Consulta(frmCR_SolicitudesFormalizacion.txtCodigoCredito, frmCR_SolicitudesFormalizacion.txtNumeroSolicitud)
 Call barra1(Me.tlbPrincipal, 2, mboton)
End If

Call RefrescaTags(Me)

End Sub

Private Sub imgBusqueda_Rapida_Click(Index As Integer)
Dim bien As Boolean

On Error GoTo CapturaError

GLOBALES.gSQLConsulta = "select cod_cuenta,descripcion from CUENTAS"
GLOBALES.gSQLColumna = "cod_cuenta"
GLOBALES.gSQLOrden = "cod_cuenta"
GLOBALES.gSQLFiltro = " and acepta_movimientos = 'S'"
Call Br(frmCR_CtaCatalogo, Index)

If GLOBALES.gSQLResulta <> "" Then
 medCuenta.Text = GLOBALES.gSQLResulta
 medCuenta.SetFocus
End If

bien = ValidaDatos(medCuenta)

Exit Sub
CapturaError:
Call ProcedimientoErrores(Me.Name)

End Sub

Private Sub medCuenta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then If ValidaDatos(medCuenta) Then MsgBox "Cuenta Verificada", vbInformation
End Sub

Private Sub SSGriddgd_DblClick()
Dim bien As Boolean
txtConcepto = ""
txtmonto = ""
medCuenta.Text = ""
lblCuenta = ""

If ssGriddgd.Rows <> 0 Then
 With ssGriddgd
  .Col = 0
  If Trim(.Columns(0).Text) <> "" Then
     Labdesembolsos.Caption = .Columns(0).Text
  End If
  .Col = 3
  txtConcepto = Trim(.Columns(3).Text)
  .Col = 4
  mblnFormato = True
  txtmonto = Format(.Columns(4).Text, "###,###,##0.00")
  .Col = 5
  medCuenta = Trim(.Columns(5).Text)
 End With
End If

bien = ValidaDatos(medCuenta)

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strSQL As String
Dim rec As ADODB.Recordset, curMontoDisponible As Currency, curMonto As Currency

'NO ES NECESARIO CONTROL DE TRANSACCIONES EN ESTA PANTALLA.
On Error GoTo CapturaError
If Button.Key <> "Cerrar" Then
Me.MousePointer = vbHourglass
End If
With tlbPrincipal.Buttons
Select Case Button.Key
  Case "insertar"
  
         'Validar Informacion AQUI
         
         .Item(1).Enabled = False 'insertar
         .Item(2).Enabled = False 'modificar
         .Item(3).Enabled = False 'borrar
         .Item(4).Enabled = True  'salvar
         .Item(5).Enabled = True  'deshacer
         .Item(6).Enabled = False 'Ayuda
         .Item(7).Enabled = False 'Salir
         Call habentrada
         Call Limpia
         Call CargaCuenta
  Case "modificar"
          If Trim(txtConcepto) = "" Then
            MsgBox "Dos clicks en registro a modificarse", vbOKOnly
            Exit Sub
          Else
         .Item(1).Enabled = False 'insertar
         .Item(2).Enabled = False 'modificar
         .Item(3).Enabled = False 'borrar
         .Item(4).Enabled = True  'salvar
         .Item(5).Enabled = True  'deshacer
         .Item(6).Enabled = False 'Ayuda
         .Item(7).Enabled = False 'Salir
         mblnmodificar = True
         Call habentrada
         Call RefrescaTags(Me)
         End If
  Case "borrar"
        If Trim(txtConcepto) = "" Then
         MsgBox "Dos clicks en registro a borrarse", vbOKOnly
         Exit Sub
        Else
        On Error GoTo CapturaError

            strSQL = "delete from desembolsos where id_desembolso=" & Trim(Labdesembolsos.Caption)
            GLOBALES.gConDatos.Execute strSQL

            strSQL = "select * from desembolsos where codigo='" & Trim(frmCR_SolicitudesFormalizacion.txtCodigoCredito.Text) & "'"
            strSQL = strSQL & " and id_solicitud=" & frmCR_SolicitudesFormalizacion.txtNumeroSolicitud.Text
            
            Call frmCR_SolicitudesFormalizacion.TotalDesRefu(curTotalDesRefu)
            lblMontoDisponible = Format(mcurMontoApr - curTotalDesRefu - mcurPrimerAbono - mcurIntHF, "standard")
            
            Set rec = New ADODB.Recordset
            With rec
              .ActiveConnection = GLOBALES.gConDatos.ConnectionString
              .CursorType = adOpenStatic
              .Source = strSQL
              .Open
              If Not .EOF And .RecordCount >= 1 Then
                Call CargaGrid(rec, ssGriddgd)
              Else
                Call CargaGrid(rec, ssGriddgd)
              End If
            End With
            Set rec = Nothing
            Call Limpia
        Call RefrescaTags(Me)
        End If
  
  Case "guardar"
         If Valida = True Then
          Call frmCR_SolicitudesFormalizacion.TotalDesRefu(curTotalDesRefu)
          curMontoDisponible = Format(mcurMontoApr - curTotalDesRefu - mcurPrimerAbono - mcurIntHF, "standard")
          curMonto = Format(txtmonto, "#######0.00")
          If (curMontoDisponible - curMonto) >= 0 Then
          .Item(1).Enabled = True   'nsertar
          .Item(2).Enabled = True   'modificar
          .Item(3).Enabled = True   'borrar
          .Item(4).Enabled = False  'salvar
          .Item(5).Enabled = False  'deshacer
          .Item(6).Enabled = True   'Ayuda
          .Item(7).Enabled = True   'Salir
          Call Inactu
          Call RefrescaGrid
          Call DeshabEntrada
          If mblnmodificar = True Then
            mblnmodificar = False
          End If
          Call frmCR_SolicitudesFormalizacion.TotalDesRefu(curTotalDesRefu)
          lblMontoDisponible = Format(mcurMontoApr - curTotalDesRefu - mcurPrimerAbono - mcurIntHF, "standard")
          Else
           MsgBox "Monto de desembolso exede el monto disponible", vbOKOnly
          End If
          'frmCR_SolicitudesFormalizacion.tlbOpciones.Buttons.Item(2).Enabled = True 'formalizar
          'frmCR_SolicitudesFormalizacion.tlbOpciones.Buttons.Item(3).Enabled = True 'anular
          Call RefrescaTags(Me)
          End If
  Case "deshacer"
         Call Limpia
         .Item(1).Enabled = True  'Insertar
         .Item(2).Enabled = True  'Modificar
         .Item(3).Enabled = True  'Borrar
         .Item(4).Enabled = False 'Salvar
         .Item(5).Enabled = False 'Deshacer
         .Item(6).Enabled = True  'Ayuda
         .Item(7).Enabled = True  'salir
          If mblnmodificar = True Then
            mblnmodificar = False
          End If
          DeshabEntrada
  Case "ayuda"
  Case "Cerrar"
   Unload Me
End Select 'button.key
End With

If Button.Key <> "Cerrar" Then
Me.MousePointer = vbDefault
End If
Exit Sub
CapturaError:
Me.MousePointer = vbDefault
Call ProcedimientoErrores(Me.Name)

End Sub


Private Sub txtmonto_Change()
If mblnFormato = False Then
 Set GLOBALES.gCajaTxt = txtmonto
 ValidaMonto
End If
End Sub

Function ValidaDatos(str As MaskEdBox) As Boolean
Dim i As Integer

On Error GoTo CapturaError
For i = Len(str.Text) To intCaracteres - 1
  str.Text = str.Text + "0"
Next i

ValidaDatos = False
With rsCuentas
 .Source = "select * from cuentas where cod_cuenta = '" & str.Text & "'"
 .CursorType = adOpenStatic
 .ActiveConnection = GLOBALES.gConDatos
 .Open
  
  
 If .EOF = True And .BOF = True Then
  MsgBox "No se encontró código de cuenta especificado, Corrijalo...", vbCritical
 Else

ValidaDatos = True
lblCuenta.Caption = !Descripcion
 
 End If
 .Close
End With

Exit Function
CapturaError:
Call ProcedimientoErrores(Me.Name)

End Function

Sub CargaLBLSDatos(str As MaskEdBox)

Dim i As Integer

On Error GoTo CapturaError
For i = Len(str.Text) To intCaracteres - 1
  str.Text = str.Text + "0"
Next i

With rsCuentas
 .Source = "select * from cuentas where cod_cuenta = '" & str.Text & "'"
 .CursorType = adOpenStatic
 .ActiveConnection = GLOBALES.gConDatos
 .Open
  
  
 If .EOF = True And .BOF = True Then
  MsgBox "No se encontró código de cuenta especificado, Corrijalo...", vbCritical
 Else
  lblCuenta.Caption = !Descripcion
 End If
 
 .Close

End With

Exit Sub
CapturaError:
 Call ProcedimientoErrores(Me.Name)
End Sub

Sub CargaCuenta()
Dim rs As New ADODB.Recordset

rs.Source = "select * from par_ahcr"
rs.ActiveConnection = GLOBALES.gConDatos
rs.CursorType = adOpenStatic
rs.Open

medCuenta.Format = GLOBALES.gstrMascara

medCuenta.Text = IIf(IsNull(rs!cr_cta_desembolso), "", rs!cr_cta_desembolso)
rs.Close

Call CargaLBLSDatos(medCuenta)

End Sub

Private Sub txtMonto_GotFocus()
mblnFormato = False
End Sub

Private Sub txtMonto_LostFocus()
 If Trim(txtmonto) <> "" Then
    mblnFormato = True
    txtmonto = Format(txtmonto, "###,###,###,##0.00")
 End If
End Sub
