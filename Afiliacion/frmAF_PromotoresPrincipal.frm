VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAF_PromotoresPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promotores"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   HelpContextID   =   1008
   Icon            =   "frmAF_PromotoresPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmAF_PromotoresPrincipal.frx":08CA
   ScaleHeight     =   5280
   ScaleWidth      =   8175
   Begin VB.TextBox txtContacto 
      Height          =   315
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   19
      ToolTipText     =   "Nombre del Promotor"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtCuenta 
      Height          =   315
      Left            =   6600
      MaxLength       =   25
      TabIndex        =   16
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      Left            =   1080
      TabIndex        =   15
      Top             =   840
      Width           =   4335
   End
   Begin VB.ComboBox cboDocumento 
      Height          =   315
      ItemData        =   "frmAF_PromotoresPrincipal.frx":0C0C
      Left            =   1080
      List            =   "frmAF_PromotoresPrincipal.frx":0C16
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtComision 
      Height          =   315
      Left            =   6600
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpFechaIngreso 
      Height          =   315
      Left            =   6600
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20316163
      CurrentDate     =   36093
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1080
      MaxLength       =   35
      TabIndex        =   4
      ToolTipText     =   "Nombre del Promotor"
      Top             =   480
      Width           =   4335
   End
   Begin VB.Frame frmEstatus 
      Caption         =   "Estatus del Promotor"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   1
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
      Begin VB.OptionButton optInactivo 
         Caption         =   "Inactivo"
         Height          =   315
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optActivo 
         Caption         =   "Activo"
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lswPromotores 
      Height          =   2655
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Object.ToolTipText     =   "Imprime el listado seleccionado"
            Object.Tag             =   "1"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFRA"
                  Object.Tag             =   "1"
                  Text            =   "Resumen de Afiliaciones"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFDA"
                  Object.Tag             =   "1"
                  Text            =   "Detalle de Afiliaciones "
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFLP"
                  Object.Tag             =   "1"
                  Text            =   "Listado de Promotores"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Consulta registros"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "cerrar"
            Object.ToolTipText     =   "Sale de esta ventana"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   5640
         Top             =   1320
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
               Picture         =   "frmAF_PromotoresPrincipal.frx":0C31
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":150D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":1DE9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":2105
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":2421
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":2CFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":3019
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":3335
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_PromotoresPrincipal.frx":3C11
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Depositar A"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Cuenta"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Banco"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Comisión"
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblConsulta 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Promotores"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   8175
   End
   Begin VB.Label Label3 
      Caption         =   "Ingreso"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmAF_PromotoresPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnEditar As Boolean
Dim mblnEOF As Boolean

Sub FormatoGrid()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla los Promotores.
'REFERENCIAS:   Ninguna
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim recPromotores As New ADODB.Recordset
Dim strSQL As String
Dim itmX As ListItem

With recPromotores
     strSQL = "Select * from Promotores order by nombre"
     .Source = strSQL
     .ActiveConnection = glogon.Conection
     .CursorType = adOpenStatic
     .Open
     
     If .EOF = True Then
        mblnEOF = True
     Else
        mblnEOF = False
     End If
     
     lswPromotores.ListItems.Clear
     Do While .EOF = False
        Set itmX = lswPromotores.ListItems.Add(, , !id_promotor)
            itmX.SubItems(1) = !Nombre
                    
            Select Case !Estado
             Case 0
              strSQL = "Inactivo"
             Case 1
              strSQL = "Activo"
            End Select
            
            itmX.SubItems(2) = strSQL
            itmX.SubItems(3) = Format(!FECHAING, "dd/mm/yyyy")
            itmX.SubItems(4) = IIf(IsNull(!Cod_Comision) = True, "", !Cod_Comision)
            
            Select Case !Tipo_Documento
             Case "CK"
              strSQL = "Cheque"
             Case "TE"
              strSQL = "Transferencia"
             Case Else
              strSQL = ""
            End Select
            
            itmX.SubItems(5) = strSQL
            itmX.SubItems(6) = fxDescribeBanco(IIf(IsNull(!cod_banco) = True, 0, !cod_banco))
            itmX.SubItems(7) = IIf(IsNull(!Cuenta_Ahorros) = True, "", !Cuenta_Ahorros)
            itmX.SubItems(8) = IIf(IsNull(!Nombre_Contacto), "", Trim(!Nombre_Contacto))
        .MoveNext
     Loop
     
     .Close
End With


End Sub

Sub Guardar()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Este Procedimiento Verifica que Los Objetos de Entrada de Datos Necesarios
'               para Guardar el Registro, contengan los datos respectivos, de ser asi,
'               procedemos a verificar si la variable mblnEditar esta en Modo de Insercion
'               o en modo de Edicion. Finalmente Guardamos el Registro.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'               FormatoGrid - (Despliega Promotores en pantalla)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strFecha As String
Dim strSQL As String

If Trim(txtComision) <> "" And Trim(cboDocumento) <> "" And Trim(cboBanco) <> "" _
   And Trim(txtContacto) <> "" _
   And Trim(txtCuenta) <> "" And Trim(txtNombre) <> "" And (optActivo(0).Value <> False Or optInactivo(1).Value <> False) Then
    strFecha = dtpFechaIngreso.Month & "/" & dtpFechaIngreso.Day & "/" & dtpFechaIngreso.Year
    If mblnEditar = True Then
       strSQL = "Update Promotores Set Nombre='" & UCase(Trim(txtNombre))
       strSQL = strSQL & "',Estado=" & IIf(optActivo(0).Value = True, 1, 0)
       strSQL = strSQL & ",FechaIng='" & strFecha & "'"
       strSQL = strSQL & ",Cod_Comision='" & Trim(txtComision) & "'"
       strSQL = strSQL & ",Tipo_Documento='" & IIf(cboDocumento = "Cheque", "CK", "TE") & "'"
       strSQL = strSQL & ",Cod_Banco=" & cboBanco.ItemData(cboBanco.ListIndex)
       strSQL = strSQL & ",Cuenta_AHORROS='" & Trim(txtCuenta) & "'"
       strSQL = strSQL & ",Nombre_Contacto='" & Trim(txtContacto) & "'"
       strSQL = strSQL & " Where Id_Promotor=" & lswPromotores.SelectedItem.Text
       glogon.Conection.Execute strSQL
       Call Bitacora("Modifica", "Modifico al Promotor " & Trim(txtNombre))
    Else
       strSQL = "Insert Into Promotores (Nombre,Estado,FechaIng,Cod_Comision,"
       strSQL = strSQL & "Tipo_Documento,Cod_Banco,Cuenta_Ahorros,"
       strSQL = strSQL & "Nombre_Contacto)"
       strSQL = strSQL & " Values('" & UCase(Trim(txtNombre)) & "',"
       strSQL = strSQL & IIf(optActivo(0).Value = True, 1, 0) & ",'"
       strSQL = strSQL & strFecha & "','"
       strSQL = strSQL & Trim(txtComision) & "','" & IIf(cboDocumento = "Cheque", "CK", "TE")
       strSQL = strSQL & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",'"
       strSQL = strSQL & Trim(txtCuenta) & "','" & Trim(txtContacto) & "')"
       
       glogon.Conection.Execute strSQL
       Call Bitacora("Registra", "Registro al Promotor " & Trim(txtNombre))
    End If
    
    Call FormatoGrid
    
    tlbPrincipal.Buttons.Item(1).Enabled = True
    tlbPrincipal.Buttons.Item(2).Enabled = True
    tlbPrincipal.Buttons.Item(3).Enabled = True
    tlbPrincipal.Buttons.Item(4).Enabled = False
    tlbPrincipal.Buttons.Item(5).Enabled = False
    tlbPrincipal.Buttons.Item(6).Enabled = True
    tlbPrincipal.Buttons.Item(7).Enabled = True
    tlbPrincipal.Buttons.Item(8).Enabled = True
    tlbPrincipal.Buttons.Item(9).Enabled = True
    
    lswPromotores.Enabled = True
    txtNombre.Enabled = False
    frmEstatus(1).Enabled = False
    dtpFechaIngreso.Enabled = False
    
    txtNombre.Text = ""
    optActivo(0).Value = False
    optInactivo(1).Value = False
    txtComision = ""
    txtComision.Enabled = False
    cboDocumento = ""
    cboDocumento.Enabled = False
    cboBanco = ""
    cboBanco.Enabled = False
    txtCuenta = ""
    txtCuenta.Enabled = False
    txtContacto = ""
    txtContacto.Enabled = False
    
    Call RefrescaTags(Me)
Else
    MsgBox "Faltan Datos", vbExclamation, "Atencion!"
End If
    
End Sub

Sub Modificar()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Este Procedimiento Habilita Todos los Objetos de Entrada de Datos .Pone la
'               Variable mblnEditar en True, lo cual Indica que estamos en Modo de Edicion
'               o Modificacion. Finalmente desplegamos en pantalla los datos originales
'               del Registro a Modificar.
'REFERENCIAS:   RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

txtNombre.Enabled = True
frmEstatus(1).Enabled = True
txtComision.Enabled = True
cboDocumento.Enabled = True
cboBanco.Enabled = True
txtCuenta.Enabled = True
txtContacto.Enabled = True


tlbPrincipal.Buttons.Item(1).Enabled = False
tlbPrincipal.Buttons.Item(2).Enabled = False
tlbPrincipal.Buttons.Item(3).Enabled = True
tlbPrincipal.Buttons.Item(4).Enabled = True
tlbPrincipal.Buttons.Item(5).Enabled = True
tlbPrincipal.Buttons.Item(6).Enabled = False
tlbPrincipal.Buttons.Item(7).Enabled = False
tlbPrincipal.Buttons.Item(8).Enabled = False
tlbPrincipal.Buttons.Item(9).Enabled = False

Call RefrescaTags(Me)
          
lswPromotores.Enabled = False
mblnEditar = True

txtNombre = Trim(lswPromotores.SelectedItem.SubItems(1))

If Trim(lswPromotores.SelectedItem.SubItems(2)) = "Activo" Then
   optActivo(0).Value = True
Else
   optInactivo(1).Value = True
End If

dtpFechaIngreso = Format(lswPromotores.SelectedItem.SubItems(3), "dd/mm/yyyy")
txtComision = Trim(lswPromotores.SelectedItem.SubItems(4))
cboDocumento = Trim(lswPromotores.SelectedItem.SubItems(5))
cboBanco = ""
txtCuenta = Trim(lswPromotores.SelectedItem.SubItems(7))
txtContacto = Trim(lswPromotores.SelectedItem.SubItems(8))

txtNombre.SetFocus
End Sub




Private Sub cboBanco_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
       txtComision.SetFocus
  Case Else
       KeyAscii = 0
End Select
End Sub


Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
     txtContacto.SetFocus
  Case Else
     KeyAscii = 0
End Select
End Sub


Private Sub dtpFechaIngreso_GotFocus()
dtpFechaIngreso.MaxDate = Format(fxFechaServidor, "dd/mm/yyyy")
End Sub

Private Sub dtpFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
       cboBanco.SetFocus
End Select
End Sub


Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verificar y establecer permisos sobre el formulario, y despliega los
'               Promotores en pantalla.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               sbToolBarIconos - (Carga los iconos para la barra de herramientas)
'               FormatoGrid - (Despliega los Promotores)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset
Dim strSQL As String

On Error GoTo CapturaError

vModulo = 1
Call Formularios(Me)

Call sbToolBarIconos(tlbPrincipal, False)

lswPromotores.ColumnHeaders.Add 1, , "ID_Promotor", 0
lswPromotores.ColumnHeaders.Add 2, , "Nombre", 4000
lswPromotores.ColumnHeaders.Add 3, , "Estado", 1500
lswPromotores.ColumnHeaders.Add 4, , "Fec.Ingreso.", 1300
lswPromotores.ColumnHeaders.Add 5, , "Comisión", 1500
lswPromotores.ColumnHeaders.Add 6, , "Documento", 1500
lswPromotores.ColumnHeaders.Add 7, , "Banco", 1500
lswPromotores.ColumnHeaders.Add 8, , "Cuenta", 1500
lswPromotores.ColumnHeaders.Add 9, , "Contacto", 0

strSQL = "Select * From Bancos where Aplica_cheques=1"
With rs
  .Open strSQL, glogon.Conection, adOpenStatic
     Do While .EOF = False
        cboBanco.AddItem Trim(!descripcion)
        cboBanco.ItemData(cboBanco.NewIndex) = !Id_banco
        .MoveNext
     Loop
  .Close
End With

Call FormatoGrid
   
If mblnEOF = True Then
    tlbPrincipal.Buttons.Item(2).Enabled = False
    tlbPrincipal.Buttons.Item(3).Enabled = False
    tlbPrincipal.Buttons.Item(6).Enabled = False
    tlbPrincipal.Buttons.Item(7).Enabled = False
    lswPromotores.Enabled = False
End If

    tlbPrincipal.Buttons.Item(4).Enabled = False
    tlbPrincipal.Buttons.Item(5).Enabled = False
    txtNombre.Enabled = False
    frmEstatus(1).Enabled = False
    dtpFechaIngreso.Enabled = False
    optActivo(0).Value = False
    optInactivo(1).Value = False
    txtComision.Enabled = False
    cboDocumento = ""
    cboDocumento.Enabled = False
    cboBanco = ""
    cboBanco.Enabled = False
    txtCuenta.Enabled = False
    txtContacto.Enabled = False
    
    Call RefrescaTags(Me)

Exit Sub
CapturaError:
   Call ProcedimientoErrores(Me.Name, Err)

End Sub




Private Sub lswPromotores_DblClick()
Call Modificar

End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Realizar el mantenimiento a los Promotores.
'REFERENCIAS:   Modificar - (Habilita Todos los Objetos de Entrada de Datos y despliega en
'               pantalla datos originales del Registro a Modificar)
'               Bitacora - (Registra movimientos sobre la Base de datos)
'               FormatoGrid - (Despliega en pantalla Promotores)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               Guardar - (Inserta o actualiza registros de Promotores)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strResp As String
Dim strSQL As String

On Error GoTo ErrorTransaccion
glogon.Conection.BeginTrans

If Button.Key <> "cerrar" Then
   Me.MousePointer = vbHourglass
End If

Select Case Button.Key
    Case "insertar"
            txtNombre.Enabled = True
            txtNombre.Text = ""
            txtNombre.SetFocus
            frmEstatus(1).Enabled = True
            dtpFechaIngreso.Enabled = True
            txtComision.Enabled = True
            txtComision = ""
            cboDocumento = ""
            cboDocumento.Enabled = True
            cboBanco = ""
            cboBanco.Enabled = True
            txtCuenta.Enabled = True
            txtCuenta = ""
            lswPromotores.Enabled = False
            txtContacto.Enabled = True
            txtContacto = ""
            mblnEditar = False
            
            dtpFechaIngreso = Format(fxFechaServidor, "dd/mm/yyyy")
            optActivo(0).Value = True
            
            tlbPrincipal.Buttons.Item(1).Enabled = False
            tlbPrincipal.Buttons.Item(2).Enabled = False
            tlbPrincipal.Buttons.Item(3).Enabled = False
            tlbPrincipal.Buttons.Item(4).Enabled = True
            tlbPrincipal.Buttons.Item(5).Enabled = True
            tlbPrincipal.Buttons.Item(6).Enabled = False
            tlbPrincipal.Buttons.Item(7).Enabled = False
            tlbPrincipal.Buttons.Item(8).Enabled = False
            tlbPrincipal.Buttons.Item(9).Enabled = False
                     
    Case "modificar"
         Call Modificar
         
    Case "borrar"
          If lswPromotores.SelectedItem.Text <> "" Then
              strResp = MsgBox("Registro Será Eliminado", vbQuestion + vbYesNo, "Confirma Eliminación?")
              If strResp = vbYes Then
                strSQL = "Delete From Promotores Where Id_Promotor=" & lswPromotores.SelectedItem.Text
                glogon.Conection.Execute strSQL
                Call Bitacora("Borra", "Elimino al Promotor " & Trim(txtNombre))
                
                txtNombre = ""
                txtNombre.Enabled = False
                frmEstatus(1).Enabled = False
                dtpFechaIngreso.Enabled = False
                optActivo(0).Value = False
                optInactivo(1).Value = False
                txtComision = ""
                txtComision.Enabled = False
                cboDocumento = ""
                cboDocumento.Enabled = False
                cboBanco = ""
                cboBanco.Enabled = False
                txtCuenta = ""
                txtCuenta.Enabled = False
                txtContacto = ""
                txtContacto.Enabled = False
                
                mblnEditar = False
                Call FormatoGrid
                
                If mblnEOF = True Then
                   tlbPrincipal.Buttons.Item(1).Enabled = True
                   tlbPrincipal.Buttons.Item(2).Enabled = False
                   tlbPrincipal.Buttons.Item(3).Enabled = False
                   tlbPrincipal.Buttons.Item(4).Enabled = False
                   tlbPrincipal.Buttons.Item(5).Enabled = False
                   tlbPrincipal.Buttons.Item(6).Enabled = False
                   tlbPrincipal.Buttons.Item(7).Enabled = False
                   tlbPrincipal.Buttons.Item(8).Enabled = True
                   tlbPrincipal.Buttons.Item(9).Enabled = True
                   lswPromotores.Enabled = False
                Else
                   tlbPrincipal.Buttons.Item(1).Enabled = True
                   tlbPrincipal.Buttons.Item(2).Enabled = True
                   tlbPrincipal.Buttons.Item(3).Enabled = True
                   tlbPrincipal.Buttons.Item(4).Enabled = False
                   tlbPrincipal.Buttons.Item(5).Enabled = False
                   tlbPrincipal.Buttons.Item(6).Enabled = True
                   tlbPrincipal.Buttons.Item(7).Enabled = True
                   tlbPrincipal.Buttons.Item(8).Enabled = True
                   tlbPrincipal.Buttons.Item(9).Enabled = True
                   lswPromotores.Enabled = True
                End If
                Call RefrescaTags(Me)
              End If
            End If
    
    Case "guardar"
         Call Guardar
    
    Case "deshacer"
            txtNombre.Text = ""
            mblnEditar = False
            txtNombre.Enabled = False
            frmEstatus(1).Enabled = False
            dtpFechaIngreso.Enabled = False
            optActivo(0).Value = False
            optInactivo(1).Value = False
            
            txtComision = ""
            txtComision.Enabled = False
            cboDocumento = ""
            cboDocumento.Enabled = False
            cboBanco = ""
            cboBanco.Enabled = False
            txtCuenta = ""
            txtCuenta.Enabled = False
            txtContacto = ""
            txtContacto.Enabled = False
            
            tlbPrincipal.Buttons.Item(1).Enabled = True
            tlbPrincipal.Buttons.Item(4).Enabled = False
            tlbPrincipal.Buttons.Item(5).Enabled = False
            tlbPrincipal.Buttons.Item(8).Enabled = True
            tlbPrincipal.Buttons.Item(9).Enabled = True
                
            If mblnEOF = True Then
               tlbPrincipal.Buttons.Item(2).Enabled = False
               tlbPrincipal.Buttons.Item(3).Enabled = False
               tlbPrincipal.Buttons.Item(6).Enabled = False
               tlbPrincipal.Buttons.Item(7).Enabled = False
               lswPromotores.Enabled = False
            Else
               tlbPrincipal.Buttons.Item(2).Enabled = True
               tlbPrincipal.Buttons.Item(3).Enabled = True
               tlbPrincipal.Buttons.Item(6).Enabled = True
               tlbPrincipal.Buttons.Item(7).Enabled = True
               lswPromotores.Enabled = True
            End If
            Call RefrescaTags(Me)

    Case "ayuda"
         MDIPrincipal.dlg.HelpContext = Me.HelpContextID
         MDIPrincipal.dlg.ShowHelp
    
    Case "cerrar"
        Unload Me
    
    Case "consultar"
'        Call bp(frmAF_PromotoresPrincipal, 3)

End Select

glogon.Conection.CommitTrans

If Button.Key <> "cerrar" Then
   Me.MousePointer = vbDefault
End If

Exit Sub
ErrorTransaccion:
 Me.MousePointer = vbDefault
 glogon.Conection.RollbackTrans
 Call ProcedimientoErrores(Me.Name, Err)

End Sub

Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Desplegar en pantalla el reporte elegido por el Usuario o el formulario
'               de reportes por promotores.
'REFERENCIAS:   ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo CapturaError
Me.MousePointer = vbHourglass

Select Case ButtonMenu.Key
Case "AFRA"
 GLOBALES.gstrReporte = "Resumen"
 frmAF_PromotoresReportes.Show vbModal

Case "AFDA"
 GLOBALES.gstrReporte = "Detalle"
 frmAF_PromotoresReportes.Show vbModal

Case "AFLP"
   With MDIPrincipal.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Afiliación"
    
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .ReportFileName = App.Path + "\Afiliacion\Reportes\ListadoPromotores.rpt"
    
    .PrintReport
   End With
   
End Select
Me.MousePointer = vbDefault

Exit Sub
CapturaError:
  Me.MousePointer = vbDefault
  Call ProcedimientoErrores(Me.Name, Err)
End Sub





Private Sub txtComision_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case vbKeyReturn
      txtCuenta.SetFocus
End Select

End Sub


Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
       cboDocumento.SetFocus
  Case 48 To 57, 8
  Case Else
    KeyAscii = 0
End Select
End Sub


Private Sub txtNombre_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case vbKeyReturn
      If dtpFechaIngreso.Enabled = True Then
         dtpFechaIngreso.SetFocus
      Else
         cboBanco.SetFocus
      End If
End Select
End Sub


