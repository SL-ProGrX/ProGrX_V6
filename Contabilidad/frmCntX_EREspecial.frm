VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_EREspecial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Confección de Estados de Resultados Especiales (Comercial)"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAccion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtDetalle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Detalle"
      Top             =   7020
      Width           =   8292
   End
   Begin VB.TextBox txtTitulo 
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
      Left            =   840
      TabIndex        =   6
      ToolTipText     =   "(F4) Nombre de la Contabilidad"
      Top             =   840
      Width           =   6135
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "(F4) Nombre de la Contabilidad"
      Top             =   480
      Width           =   5175
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "(F4) Código de la Contabilidad (Auto)"
      Top             =   480
      Width           =   975
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5652
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cuentas"
      TabPicture(0)   =   "frmCntX_EREspecial.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgExplorer"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ArbolExp"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboGrupo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.ComboBox cboGrupo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   3015
      End
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   4680
         Left            =   120
         TabIndex        =   1
         Top             =   888
         Width           =   8052
         _ExtentX        =   14208
         _ExtentY        =   8255
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   3
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imgExplorer"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imgExplorer 
         Left            =   6240
         Top             =   4320
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
               Picture         =   "frmCntX_EREspecial.frx":001C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":08F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":11D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":14F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":180C
               Key             =   "imgRoot"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":1B28
               Key             =   "imgFolder"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":1E44
               Key             =   "imgOpcion"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":2160
               Key             =   "imgUsuario"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":2A3C
               Key             =   "imgGrupo"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":3318
               Key             =   "imgAsientos"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_EREspecial.frx":3634
               Key             =   "imgCuentas"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Acción"
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
         Left            =   3960
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton cmdActualizaGrupo 
      Height          =   732
      Left            =   7080
      TabIndex        =   13
      Top             =   480
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Actualiza Grupo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCntX_EREspecial.frx":3950
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
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
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ER"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmCntX_EREspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vEdita As Boolean, vCodigo As Long, vTipoBusca As String
Dim vBuscaTree As Boolean


Private Sub cboAccion_Click()
If Mid(cboAccion.Text, 1, 2) <> "" Then
  vBuscaTree = True
  Call sbRefrescaArbol
End If
End Sub

Private Sub cboGrupo_Click()

If Mid(cboGrupo.Text, 1, 2) <> "" Then
  vBuscaTree = True
  Call sbRefrescaArbol
End If

End Sub

Private Sub cmdActualizaGrupo_Click()
If vEdita Then Call sbGuarda_Cuentas
End Sub

Private Sub Form_Load()

On Error GoTo vError

Set Me.Icon = frmContenedor.Icon
vBuscaTree = False
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call sbRefrescaArbol

 Call Formularios(Me)
 Call RefrescaTags(Me)
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Sub sbLimpiaPantalla()
vTipoBusca = "D"
vCodigo = 0
txtCodigo = ""
txtNombre = ""
txtCodigo.Enabled = True
txtDetalle = ""
txtTitulo = ""

vBuscaTree = False
With cboAccion
  .Clear
  .AddItem "01 - Definición"
  .AddItem "02 - Efecto Positivo"
  .AddItem "03 - Efecto Negativo"
  .Text = "01 - Definición"
End With

With cboGrupo
  .Clear
  .AddItem "01 - Ventas"
  .AddItem "02 - Costo Mercaderia Vendida"
  .AddItem "03 - Gastos"
  .AddItem "04 - Ingresos"
  .Text = "01 - Ventas"
End With

vBuscaTree = True


End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
    Case "CONSULTAR"
       If vTipoBusca = "D" Then
         gBusquedas.Columna = "descripcion"
         gBusquedas.Orden = "descripcion"
       Else
         gBusquedas.Columna = "cod_er_especial"
         gBusquedas.Orden = "cod_er_especial"
       End If
       gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
       gBusquedas.Consulta = "select cod_er_especial,descripcion,titulo from er_especial"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    
    Case "REPORTES"
      If vCodigo > 0 Then
       strSQL = "{ER_ESPECIAL_DK.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
              & " AND {ER_ESPECIAL_DK.COD_ER_ESPECIAL} = " & vCodigo
       Call sbCntX_Reportes("CATALOGOERESPECIAL", strSQL, "ER.ESPECIAL : " & txtNombre)
      Else
        MsgBox "Debe de Seleccionar un ER ESPECIAL Antes de Imprimir...", vbInformation
      End If
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    
    Case "CERRAR"
      UnLoad Me
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from er_especial where cod_er_especial = " & lngCodigo _
       & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_er_especial
  'llenar datos en pantalla
  txtCodigo = rs!cod_er_especial
  txtNombre = Trim(rs!Descripcion & "")
  txtTitulo = Trim(rs!titulo & "")
  txtDetalle = "US-CREA:" & rs!user_crea & " FECHA-CREA:" & Format(rs!fecha_crea, "yyyy/mm/dd")
  vBuscaTree = True
  Call sbRefrescaArbol
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDecimal
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Reporte no es válido ..."
If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Titulo del Reporte no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Function fxExisteCuenta(xCuenta As String) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset, vOperacion As String

If vCodigo = 0 Then
  fxExisteCuenta = False
  Exit Function
End If


Select Case Mid(cboAccion.Text, 1, 2)
  Case "01"
    vOperacion = "D"
  Case "02"
    vOperacion = "P"
  Case "03"
    vOperacion = "N"
End Select


strSQL = "select isnull(count(*),0) as Existe from er_especial_dk" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_cuenta = '" _
       & xCuenta & "' and cod_er_especial = " & vCodigo _
       & " and bloque = '" & Mid(cboGrupo.Text, 1, 2) & "' and operacion = '" & vOperacion & "'"
Call OpenRecordSet(rsX, strSQL, 0)
fxExisteCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function

Private Sub sbGuarda_Cuentas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long, xNode As Node, vOperacion As String

On Error GoTo vError

frmCntX_Procesos.Show

frmCntX_Procesos.Caption = "Actualizando CntX_Cuentas..."
frmCntX_Procesos.prgBar.Max = ArbolExp.Nodes.Count
frmCntX_Procesos.prgBar.Value = 1
frmCntX_Procesos.Refresh

Select Case Mid(cboAccion.Text, 1, 2)
  Case "01"
    vOperacion = "D"
  Case "02"
    vOperacion = "P"
  Case "03"
    vOperacion = "N"
End Select
 
strSQL = "delete er_especial_dk where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_er_especial = " & vCodigo & " and bloque = '" _
       & Mid(cboGrupo.Text, 1, 2) & "' and operacion = '" & vOperacion & "'"
Call ConectionExecute(strSQL, 0)

strSQL = ""
'Carga nuevos Datos
For lng = 1 To ArbolExp.Nodes.Count
  If ArbolExp.Nodes(lng).Checked Then
    If Right(ArbolExp.Nodes(lng).Key, 1) = "C" Then
        If Not fxExisteCuenta(fxIndiceCodigo(ArbolExp.Nodes(lng).Key)) Then
           strSQL = strSQL & Space(10) & "Insert er_especial_dk(cod_contabilidad,cod_er_especial,cod_cuenta,bloque,operacion) values(" _
                  & gCntX_Parametros.CodigoConta & "," & vCodigo & ",'" _
                  & Trim(fxIndiceCodigo(ArbolExp.Nodes(lng).Key)) & "','" _
                  & Mid(cboGrupo.Text, 1, 2) & "','" & vOperacion & "')"
        End If
    End If
  End If

  frmCntX_Procesos.prgBar.Value = lng
Next lng

If Len(strSQL) > 0 Then
           Call ConectionExecute(strSQL, 0)
End If

'Verificar Cadena de Madres Aqui

'For lng = 1 To 5 'Recorre 5 veces para Actualizar todos los niveles
'    strSQL = "select A.cod_cuenta,C.acepta_movimientos,C.cuenta_madre,C.nivel" _
'           & " from er_especial_dk A inner join CntX_Cuentas C on A.cod_cuenta = C.cod_cuenta" _
'           & " And A.cod_contabilidad = C.cod_contabilidad" _
'           & " where A.cod_er_especial = " & vCodigo & " and A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
'           & " order by A.cod_cuenta desc"
'    Call OpenRecordSet(rs, strSQL, 0)
'
'    frmCntX_Procesos.Caption = "Vericando Pertenencia de CntX_Cuentas Espere..."
'    frmCntX_Procesos.prgBar.Value = 1
'    frmCntX_Procesos.prgBar.Max = rs.RecordCount + 1
'    frmCntX_Procesos.Refresh
'    Do While Not rs.EOF
'      If Not fxExisteCuenta(rs!cuenta_madre) And rs!nivel > 1 Then
'        strSQL = "Insert er_especial_dk(cod_contabilidad,cod_er_especial,cod_cuenta,bloque,operacion) values(" _
'               & gCntX_Parametros.CodigoConta & "," & vCodigo & ",'" _
'               & Trim(rs!cuenta_madre) & "','" & Mid(cboGrupo.Text, 1, 2) & "','" _
'               & vOperacion & "')"
'        Call ConectionExecute(strSQL, 0)
'      End If
'      frmCntX_Procesos.prgBar.Value = frmCntX_Procesos.prgBar.Value + 1
'      rs.MoveNext
'    Loop
'    rs.Close
'Next lng

UnLoad frmCntX_Procesos
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  'Verificar si cambio cedula o codigo para actualización en cascada
  strSQL = "update er_especial set descripcion = '" & UCase(Trim(txtNombre)) & "'" _
         & ",titulo = '" & Trim(txtTitulo) & "'" _
         & " where cod_er_especial = " & vCodigo _
         & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
  Call ConectionExecute(strSQL, 0)
  Call Bitacora("Modifica", "ER ESPECIAL: " & vCodigo & " EMP: " & gCntX_Parametros.CodigoConta)

Else
   strSQL = "select isnull(max(cod_er_especial),0) as ultimo from er_especial " _
          & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
   Call OpenRecordSet(rs, strSQL, 0)
     txtCodigo = rs!ultimo + 1
     vCodigo = txtCodigo
   rs.Close
   
   strSQL = "insert into er_especial(cod_contabilidad,cod_er_especial,descripcion,titulo,user_crea,fecha_crea) values(" _
          & gCntX_Parametros.CodigoConta & "," & vCodigo & ",'" & Trim(UCase(txtNombre)) & "','" _
          & Trim(txtTitulo) & "','" & glogon.Usuario & "','" & Format(fxFechaServidor, "yyyy/mm/dd") & "')"
   Call ConectionExecute(strSQL, 0)
    
   Call Bitacora("Registra", "ER ESPECIAL: " & vCodigo & " EMP: " & gCntX_Parametros.CodigoConta)
    
   txtCodigo.Enabled = True
 
End If

'Actualizar Aqui las CntX_Cuentas
Call sbGuarda_Cuentas

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete er_especial_dk where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
         & " and cod_er_especial = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete er_especial where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
         & " and cod_er_especial = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & gCntX_Parametros.CodigoConta)

  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_GotFocus()
 vTipoBusca = "C"
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtNombre_GotFocus()
 vTipoBusca = "D"
End Sub

Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

Me.MousePointer = vbHourglass

If Not vBuscaTree Then
  ArbolExp.Nodes.Clear
  Exit Sub
End If

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "CntX_Cuentas", "CntX_Cuentas", "imgRoot")
  'Crear Arbol Inicial
  
  strSQL = "select tipo_cuenta,Descripcion from CntX_Tipos_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta
  rs.Open strSQL, glogon.Conection, adOpenForwardOnly
  Do While Not rs.EOF
    Call sbCreaNodos("CntX_Cuentas", rs!Descripcion, "imgCntX_Cuentas", True, "0x0" & rs!tipo_cuenta & "T")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With

Me.MousePointer = vbDefault

End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "CntX_Cuentas" Then

Select Case Right(Node.Key, 1)
        
    Case "T" 'Tipos de CntX_Cuentas
    
        strSQL = "select cod_cuenta,descripcion from CntX_Cuentas where cuenta_madre = ''" _
               & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and tipo_cuenta = '" & fxIndiceCodigo(Node.Key) & "'"
        rs.Open strSQL, glogon.Conection, adOpenForwardOnly
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", True, "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C", True)
          rs.MoveNext
        Loop
        rs.Close
    
    Case Else 'SubCntX_Cuentas
    
        strSQL = "select cod_cuenta,descripcion from CntX_Cuentas where cuenta_madre = '" & fxCntX_CuentaFormato(False, fxIndiceCodigo(Node.Key)) _
               & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
        rs.Open strSQL, glogon.Conection, adOpenForwardOnly
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", True, "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C", True)
          rs.MoveNext
        Loop
        rs.Close
End Select

End If

End Sub


Private Sub sbCreaNodos(vPadre As String, vTexto As String _
    , vImagen As String, vExpand As Boolean, Optional xkey As String = "N", Optional xCuenta As Boolean = False)
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
    If xCuenta Then
      nodX.Checked = fxExisteCuenta(fxIndiceCodigo(nodX.Key))
    End If
    
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTitulo.SetFocus
End Sub
