VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPosPlanOfertas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Ofertas y Descuentos"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   9105
   Begin VB.TextBox txtDescuento 
      Alignment       =   1  'Right Justify
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
      Left            =   7680
      TabIndex        =   25
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cbo 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   840
      Width           =   7575
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   1320
      TabIndex        =   17
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtDescripcion 
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
      Left            =   2520
      TabIndex        =   16
      Top             =   480
      Width           =   6375
   End
   Begin VB.Frame fraActivacion 
      Caption         =   "Activar Oferta"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   7572
      Begin VB.CheckBox chkLunes 
         Caption         =   "Lunes"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkMartes 
         Caption         =   "Martes"
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
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkMiercoles 
         Caption         =   "Miércoles"
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkJueves 
         Caption         =   "Jueves"
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
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkViernes 
         Caption         =   "Viernes"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkSabados 
         Caption         =   "Sábados"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkDomingos 
         Caption         =   "Domingos"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   312
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   294649859
         CurrentDate     =   37682
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   312
         Left            =   3360
         TabIndex        =   9
         Top             =   720
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   294649859
         CurrentDate     =   37682
      End
      Begin MSComCtl2.DTPicker dtpHoraInicio 
         Height          =   312
         Left            =   3360
         TabIndex        =   10
         Top             =   1200
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   294649858
         CurrentDate     =   37682
      End
      Begin MSComCtl2.DTPicker dtpHoraCorte 
         Height          =   312
         Left            =   3360
         TabIndex        =   11
         Top             =   1560
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   294649858
         CurrentDate     =   37682
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   1680
         X2              =   1680
         Y1              =   360
         Y2              =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   1920
         TabIndex        =   14
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Hora Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   13
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Hora Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   1920
         TabIndex        =   12
         Top             =   1560
         Width           =   972
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   1005
      ButtonWidth     =   487
      ButtonHeight    =   466
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Descuento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   6480
      TabIndex        =   24
      Top             =   1560
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Linea"
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
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
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
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Oferta"
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
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmPosPlanOfertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 vModulo = 32

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla


 strSQL = "select cod_ProdClas,descripcion from pv_Prod_Clasifica"
 Call OpenRecordSet(rs, strSQL)
 Do While Not rs.EOF
  cbo.AddItem rs!Descripcion
  cbo.ItemData(cbo.NewIndex) = rs!cod_prodclas
  rs.MoveNext
 Loop
 If rs.RecordCount > 0 Then
   rs.MoveFirst
   cbo.Text = rs!Descripcion
 End If
 rs.Close


 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = 0
txtCodigo = ""

txtDescripcion = ""
txtNotas = ""

txtCodigo.Enabled = True


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpHoraInicio.Value = dtpInicio.Value
dtpHoraCorte.Value = dtpCorte.Value

chkLunes.Value = vbUnchecked
chkMartes.Value = vbUnchecked
chkMiercoles.Value = vbUnchecked
chkJueves.Value = vbUnchecked
chkViernes.Value = vbUnchecked
chkSabados.Value = vbUnchecked
chkDomingos.Value = vbUnchecked

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
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
'       gBusquedas.Columna = "descripcion"
'       gBusquedas.Orden = "descripcion"
'       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select O.*,L.descripcion as Linea" _
       & " from pv_PlanOfertas O inner join pv_Prod_Clasifica L on O.cod_prodClas = L.cod_prodclas" _
       & " where O.cod_oferta = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_oferta
  txtCodigo = rs!cod_oferta
  
  txtDescripcion = rs!Descripcion
  txtNotas = rs!Notas & ""
  
  cbo.Text = rs!Linea
  
  txtDescuento = rs!descuento
  
  dtpInicio.Value = rs!Fecha_Inicio
  dtpCorte.Value = rs!Fecha_Corte
  dtpHoraInicio.Value = rs!frecuencia_horai
  dtpHoraCorte.Value = rs!frecuencia_horac
  
  chkLunes.Value = rs!frecuencia_lunes
  chkMartes.Value = rs!frecuencia_martes
  chkMiercoles.Value = rs!frecuencia_miercoles
  chkJueves.Value = rs!frecuencia_jueves
  chkViernes.Value = rs!frecuencia_viernes
  chkSabados.Value = rs!frecuencia_sabado
  chkDomingos.Value = rs!frecuencia_domingo

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer
Dim curCantidad As Currency
Dim rs As New ADODB.Recordset

On Error GoTo vError


If vEdita Then
    strSQL = "update pv_PlanOfertas set descripcion = '" & UCase(txtDescripcion) _
           & "',notas = '" & txtNotas & "',user_modifica = '" & glogon.Usuario _
           & "',fecha_modifica = dbo.MyGetdate(),fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',frecuencia_horai = '" _
           & Format(dtpHoraInicio.Value, "hh:mm:ss") & "',frecuencia_horac = '" & Format(dtpHoraCorte.Value, "hh:mm:ss") _
           & "',frecuencia_lunes = " & chkLunes.Value & ",frecuencia_martes = " & chkMartes.Value _
           & ",frecuencia_miercoles = " & chkMiercoles.Value & ",frecuencia_jueves = " & chkJueves.Value _
           & ",frecuencia_viernes = " & chkViernes.Value & ",frecuencia_sabado = " & chkSabados.Value _
           & ",frecuencia_domingo = " & chkDomingos.Value & ",cod_prodClas = " & cbo.ItemData(cbo.ListIndex) _
           & ",descuento = " & CCur(txtDescuento) _
           & " where cod_oferta = " & vCodigo
   Call ConectionExecute(strSQL)

   Call Bitacora("Modifica", "Plan de Oferta: " & vCodigo)

Else
    strSQL = "select isnull(max(cod_oferta),0) + 1 as Oferta from pv_PlanOfertas"
    Call OpenRecordSet(rs, strSQL)
     vCodigo = rs!oferta
    rs.Close
    txtCodigo = vCodigo
    
    strSQL = "insert pv_PlanOfertas(cod_oferta,descripcion,fecha_crea,user_crea,notas" _
           & ",fecha_inicio,fecha_corte,frecuencia_horai,frecuencia_horac" _
           & ",frecuencia_lunes,frecuencia_martes,frecuencia_miercoles,frecuencia_jueves" _
           & ",frecuencia_viernes,frecuencia_sabado,frecuencia_domingo,cod_ProdClas,Descuento)" _
           & " values(" & vCodigo & ",'" & UCase(txtDescripcion) & "',dbo.MyGetdate(),'" _
           & glogon.Usuario & "','" & txtNotas & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & Format(dtpHoraInicio.Value, "hh:mm:ss") _
           & "','" & Format(dtpHoraCorte.Value, "hh:mm:ss") & "'," & chkLunes.Value _
           & "," & chkMartes.Value & "," & chkMiercoles.Value & "," & chkJueves.Value _
           & "," & chkViernes.Value & "," & chkSabados.Value & "," & chkDomingos.Value _
           & "," & cbo.ItemData(cbo.ListIndex) & "," & CCur(txtDescuento) & ")"
    Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Plan de Oferta: " & vCodigo)
End If

txtCodigo.Enabled = True

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
'  Call ConectionExecute(strSQL)

'  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & vParametros.CodigoEmpresa)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Integer, vSQL As String

vSQL = ""

Select Case UCase(ButtonMenu.Key)
  Case "REPBOLETA"
     
     i = MsgBox("Desea visualizar solo el paquete Actual", vbYesNo)
     If i = vbYes Then vSQL = "{PV_PAQUETES.COD_PAQUETE} = " & txtCodigo

     Call sbInvReportes("PaquetesBoleta", "Boleta de Paquetes", "", vSQL)

  Case "REPLISTADOGENERAL"
     Call sbInvReportes("PaquetesListado", "PAQUETES", "Listado General", vSQL)

End Select


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_oferta"
  gBusquedas.Orden = "cod_oferta"
  gBusquedas.Consulta = "select cod_oferta,descripcion,notas from pv_planOfertas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub

