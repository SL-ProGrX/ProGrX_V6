VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmCntX_AreaDefinicion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Areas de Trabajo"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkActiva 
      Appearance      =   0  'Flat
      Caption         =   "Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6840
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      Width           =   855
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6492
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   8172
      _ExtentX        =   14420
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
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
      TabPicture(0)   =   "frmCntX_AreaDefinicion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgExplorer"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ArbolExp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Unidades / Centro de Costos"
      TabPicture(1)   =   "frmCntX_AreaDefinicion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TreeUnidades"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   6000
         Left            =   120
         TabIndex        =   5
         Top             =   408
         Width           =   7932
         _ExtentX        =   13996
         _ExtentY        =   10583
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
         Left            =   6120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_AreaDefinicion.frx":0038
               Key             =   "imgRoot"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_AreaDefinicion.frx":0145
               Key             =   "imgFolder"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_AreaDefinicion.frx":0261
               Key             =   "imgCuentas"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_AreaDefinicion.frx":037F
               Key             =   "imgAsientos"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeUnidades 
         Height          =   6000
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   7932
         _ExtentX        =   13996
         _ExtentY        =   10583
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
      Left            =   920
      TabIndex        =   2
      ToolTipText     =   "(F4) Código de la Contabilidad (Auto)"
      Top             =   480
      Width           =   855
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
      TabIndex        =   1
      ToolTipText     =   "(F4) Nombre de la Contabilidad"
      Top             =   480
      Width           =   4935
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8280
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmCntX_AreaDefinicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vEdita As Boolean, vCodigo As Long, vTipoBusca As String

Private Sub ArbolExp_NodeCheck(ByVal Node As MSComctlLib.Node)
   Dim n As Integer

On Error GoTo vError:

Node.Child.Checked = Node.Checked

n = Node.Child.FirstSibling.Index
While n <> Node.LastSibling.Index
   n = ArbolExp.Nodes(n).Next.Index
   ArbolExp.Nodes(n).Checked = Node.Checked
Wend

vError:
End Sub

Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
vModulo = 20

Set Me.Icon = frmContenedor.Icon
 
Call sbToolBarIconos(tlb)

If gCntX_Arbol.ArbolActivo Then
  Call sbConsulta(Val(gCntX_Arbol.AsientoNumr))
Else
    vEdita = False
    Call sbLimpiaPantalla
    Call sbRefrescaArbol
    Call sbRefrescaArbolUnidades
    Call sbToolBar(tlb, "activo")
End If

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaPantalla()
vTipoBusca = "D"
vCodigo = 0
txtCodigo = ""
txtNombre = ""
txtCodigo.Enabled = True

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
         gBusquedas.Columna = "cod_area"
         gBusquedas.Orden = "cod_area"
       End If
       gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
       gBusquedas.Consulta = "select cod_area,descripcion from CntX_Area_Definicion"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    
    Case "REPORTES"
      If vCodigo > 0 Then
       strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
              & " AND {CntX_Area_Cuentas.COD_AREA} = " & vCodigo
       Call sbCntX_Reportes("CATALOGOAREAS", strSQL, txtNombre)
      Else
        MsgBox "Debe de Seleccionar un Area de Trabajo Antes de Imprimir...", vbInformation
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

strSQL = "select * from CntX_Area_Definicion where cod_area = " & lngCodigo _
       & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_area
  
  txtCodigo = rs!cod_area
  txtNombre = rs!Descripcion & ""
  chkActiva.Value = rs!activa
  
  Call sbRefrescaArbol
  Call sbRefrescaArbolUnidades

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

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

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del área no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Function fxExisteCuenta(xCuenta As String) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

If vCodigo = 0 Then
  fxExisteCuenta = False
  Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from CntX_Area_Cuentas" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_cuenta = '" _
       & xCuenta & "' and cod_area = " & vCodigo
Call OpenRecordSet(rsX, strSQL, 0)
fxExisteCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function

Private Sub sbGuardaCuentas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long, xNode As Node

On Error GoTo vError

frmCntX_Procesos.Show

frmCntX_Procesos.Caption = "Actualizando Cuentas..."
frmCntX_Procesos.prgBar.Max = ArbolExp.Nodes.Count
frmCntX_Procesos.prgBar.Value = 1
frmCntX_Procesos.Refresh


'Limpia Información
strSQL = "delete CntX_Area_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_area = " & vCodigo

strSQL = strSQL & Space(10) & "delete CntX_Area_Reportes where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_area = " & vCodigo
Call ConectionExecute(strSQL, 0)


'Carga nuevos Datos

strSQL = ""

For lng = 1 To ArbolExp.Nodes.Count
  If ArbolExp.Nodes(lng).Checked Then
    If Right(ArbolExp.Nodes(lng).Key, 1) = "C" Then
        If Not fxExisteCuenta(fxIndiceCodigo(ArbolExp.Nodes(lng).Key)) Then
           strSQL = strSQL & Space(10) & "Insert CntX_Area_Cuentas(cod_contabilidad,cod_area,cod_cuenta) values(" _
                  & gCntX_Parametros.CodigoConta & "," & vCodigo & ",'" _
                  & Trim(fxIndiceCodigo(ArbolExp.Nodes(lng).Key)) & "')"
        End If
    End If
  End If


    If strSQL <> "" Then
        Call ConectionExecute(strSQL, 0)
    End If

  frmCntX_Procesos.prgBar.Value = lng
Next lng

'Verificar Cadena de Madres Aqui

For lng = 1 To 5 'Recorre 5 veces para Actualizar todos los niveles
    strSQL = "select A.cod_cuenta,C.acepta_movimientos,C.cuenta_madre,C.nivel" _
           & " from CntX_Area_Cuentas A inner join CntX_Cuentas C on A.cod_cuenta = C.cod_cuenta" _
           & " And A.cod_contabilidad = C.cod_contabilidad" _
           & " where A.cod_area = " & vCodigo & " and A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " order by A.cod_cuenta desc"
    Call OpenRecordSet(rs, strSQL, 0)
    
    frmCntX_Procesos.Caption = "Vericando Pertenencia de CntX_Cuentas Espere..."
    frmCntX_Procesos.prgBar.Value = 1
    frmCntX_Procesos.prgBar.Max = rs.RecordCount + 1
    frmCntX_Procesos.Refresh
    
    
    strSQL = ""
    Do While Not rs.EOF
      If Not fxExisteCuenta(rs!cuenta_madre) And rs!nivel > 1 Then
        strSQL = strSQL & Space(10) & "Insert CntX_Area_Cuentas(cod_contabilidad,cod_area,cod_cuenta) values(" _
               & gCntX_Parametros.CodigoConta & "," & vCodigo & ",'" _
               & Trim(rs!cuenta_madre) & "')"
      End If
      frmCntX_Procesos.prgBar.Value = frmCntX_Procesos.prgBar.Value + 1
      rs.MoveNext
    Loop
    rs.Close
    
    If strSQL <> "" Then
        Call ConectionExecute(strSQL, 0)
    End If
Next lng

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
  strSQL = "update CntX_Area_Definicion set descripcion = '" & UCase(Trim(txtNombre)) & "',activa = " & chkActiva.Value _
         & " where cod_area = " & vCodigo & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Modifica", "Area Trabajo: " & vCodigo & " Conta." & gCntX_Parametros.CodigoConta)

Else
   strSQL = "select isnull(max(cod_area),0) as ultimo from CntX_Area_Definicion " _
          & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
   Call OpenRecordSet(rs, strSQL, 0)
     txtCodigo = rs!ultimo + 1
     vCodigo = txtCodigo
   rs.Close
   
   strSQL = "insert into CntX_Area_Definicion(cod_contabilidad,cod_area,descripcion,activa) values(" _
          & gCntX_Parametros.CodigoConta & "," & vCodigo & ",'" & Trim(UCase(txtNombre)) & "'," & chkActiva.Value & ")"
   Call ConectionExecute(strSQL, 0)
    
   Call Bitacora("Registra", "Area Trabajo: " & vCodigo & " Conta." & gCntX_Parametros.CodigoConta)
    
   txtCodigo.Enabled = True
 
End If

'Actualizar Aqui las CntX_Cuentas
Call sbGuardaCuentas

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
  strSQL = "delete CntX_Area_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
         & " and cod_area = " & vCodigo
  
  strSQL = strSQL & Space(10) & "delete CntX_Area_Reportes where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
         & " and cod_area = " & vCodigo
  
  strSQL = strSQL & Space(10) & "delete CntX_Area_Definicion where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
         & " and cod_area = " & vCodigo
  
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Elimina", "Area de Trabajo: " & vCodigo & " Conta." & gCntX_Parametros.CodigoConta)

  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub TreeUnidades_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then TreeUnidades.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "Unidades" Then

Select Case Right(Node.Key, 1)
        
    Case "U"
    
        strSQL = "select cod_centro_costo,descripcion from CntX_Centro_Costos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and cod_centro_costo in(select cod_centro_costo from cntX_unidades_cc where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and cod_unidad = '" & fxIndiceCodigo(Node.Key) & "')"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
          Call sbCreaNodosUnidad(Node.Key, rs!Descripcion, "imgFolder", False, "0x0" & fxIndiceCodigo(Node.Key) & "-" & rs!cod_centro_costo & "C", False)
          rs.MoveNext
        Loop
        rs.Close
    
    Case Else 'SubCuentas
    
End Select

End If

End Sub

Private Sub txtCodigo_GotFocus()
 vTipoBusca = "C"
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
If IsNumeric(txtCodigo.Text) Then
 Call sbConsulta(txtCodigo)
End If
End Sub

Private Sub txtNombre_GotFocus()
 vTipoBusca = "D"
End Sub

Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Cuentas", "Cuentas", "imgRoot")
  'Crear Arbol Inicial
  
  strSQL = "select tipo_cuenta,Descripcion from CntX_Tipos_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
    Call sbCreaNodos("Cuentas", rs!Descripcion, "imgCuentas", True, "0x0" & rs!tipo_cuenta & "T")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With

End Sub


Private Sub sbRefrescaArbolUnidades()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

With TreeUnidades
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Unidades", "Unidades", "imgRoot")
  'Crear Arbol Inicial
  
  strSQL = "select cod_unidad,Descripcion from CntX_Unidades where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
         & " and activa = 1"
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
    Call sbCreaNodosUnidad("Unidades", rs!Descripcion, "imgCuentas", True, "0x0" & rs!cod_unidad & "U")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With

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

If Node.Text <> "Cuentas" Then

Select Case Right(Node.Key, 1)
        
    Case "T" 'Tipos de CntX_Cuentas
    
        strSQL = "select cod_cuenta,descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = ''" _
               & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and tipo_cuenta = '" & fxIndiceCodigo(Node.Key) & "'"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder" _
                    , IIf((rs!Acepta_movimientos = 0), True, False), "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C", True)
          rs.MoveNext
        Loop
        rs.Close
    
    Case Else 'SubCuentas
    
        strSQL = "select cod_cuenta,descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = '" & fxCntX_CuentaFormato(False, fxIndiceCodigo(Node.Key)) _
               & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder" _
                    , IIf((rs!Acepta_movimientos = 0), True, False), "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C", True)
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


Private Sub sbCreaNodosUnidad(vPadre As String, vTexto As String _
    , vImagen As String, vExpand As Boolean, Optional xkey As String = "N", Optional xCuenta As Boolean = False)
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = TreeUnidades.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
'    If xCuenta Then
'      nodX.Checked = fxExisteCuenta(fxIndiceCodigo(nodX.Key))
'    End If
    
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub




