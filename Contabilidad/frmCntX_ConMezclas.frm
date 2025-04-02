VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmCntX_ConMezclas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mezclas de Consolidados"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtDescConsolidacion 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   1125
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrición"
         Object.Width           =   7479
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
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
      BorderStyle     =   1
   End
   Begin VB.Label label1 
      Caption         =   "Cons. Base"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mezclar Consolidaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   6855
   End
End
Attribute VB_Name = "frmCntX_ConMezclas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vTipoBusca As String
Dim vBusca As Boolean

Private Sub Form_Load()

Set Me.Icon = frmContenedor.Icon
 
 vEdita = True
 
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaPantalla()
vBusca = True
vTipoBusca = "D"
vCodigo = 0
txtCodigo = ""
txtDescConsolidacion = ""

txtCodigo.Enabled = True

lsw.ListItems.Clear

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = True
      
      txtDescConsolidacion.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescConsolidacion.SetFocus
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
         gBusquedas.Columna = "cod_consolida"
         gBusquedas.Orden = "cod_consolida"
       End If
       gBusquedas.Filtro = ""
       gBusquedas.Consulta = "select cod_consolida,descripcion from CNTX_CONSOLIDA_DEFINICION"
       frmBusquedas.Show vbModal

       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtDescConsolidacion = IIf((gBusquedas.Resultado2 = ""), 0, gBusquedas.Resultado2)
       txtDescConsolidacion.SetFocus
    
    Case "REPORTES"
      If vCodigo > 0 Then
'       strSQL = "{CUENTAS.COD_CONTABILIDAD} = " & gCntX_Parametros.CodigoConta _
'              & " AND {AREA_CUENTAS.COD_AREA} = " & vCodigo
'       Call sbReportes("CATALOGOAREAS", strSQL, txtNombre)
      Else
        MsgBox "Debe de Seleccionar un Area de Trabajo Antes de Imprimir...", vbInformation
      End If
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim rsTmp As New ADODB.Recordset, itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & lngCodigo
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vBusca = False
  vCodigo = rs!COD_CONSOLIDA
  'llenar datos en pantalla
  txtCodigo = rs!COD_CONSOLIDA
  txtDescConsolidacion = rs!Descripcion & ""
  
  'Busca Otros Consolidados con Mascara Semejante
  '1. Las marcadas
  strSQL = "select * from CNTX_CONSOLIDA_DEFINICION where cod_consolida in(" _
         & "select M.con_consolida_aux" _
         & " from con_mezclas M where M.cod_consolida = " & lngCodigo & ")"
  rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rsTmp.EOF
    If rsTmp!COD_CONSOLIDA <> lngCodigo Then
      Set itmX = lsw.ListItems.Add(, , rsTmp!COD_CONSOLIDA)
          itmX.SubItems(1) = rsTmp!Descripcion & ""
          itmX.Checked = True
    End If
    rsTmp.MoveNext
  Loop
  rsTmp.Close
  
  '1. Las no marcadas
  strSQL = "select * from CNTX_CONSOLIDA_DEFINICION where cod_consolida not in(" _
         & "select M.con_consolida_aux" _
         & " from con_mezclas M where M.cod_consolida = " & lngCodigo & ")"
  rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rsTmp.EOF
    If rsTmp!COD_CONSOLIDA <> lngCodigo Then
      Set itmX = lsw.ListItems.Add(, , rsTmp!COD_CONSOLIDA)
          itmX.SubItems(1) = rsTmp!Descripcion & ""
    End If
    rsTmp.MoveNext
  Loop
  rsTmp.Close
  
  
  
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

If txtDescConsolidacion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion de la Consolidacion no es valida ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long


On Error GoTo vError

If vEdita Then
  'Nada se mueve a puro detalle
  Call Bitacora("Modifica", "Consolidacion Mezcla : " & vCodigo)

Else
  Call Bitacora("Registra", "Consolidación Mezcla: " & vCodigo)
    
   txtCodigo.Enabled = True
 
End If

'Actualizar Aqui Guarda las CNTX_CONTABILIDADES Asociadas
strSQL = "delete con_mezclas where cod_consolida = " & vCodigo
glogon.Conection.Execute strSQL

For lng = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(lng).Checked Then
    strSQL = "insert into con_mezclas(cod_consolida,con_consolida_aux) values(" _
           & vCodigo & "," & lsw.ListItems.Item(lng).Text & ")"
    glogon.Conection.Execute strSQL
  End If
Next lng

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
  
  strSQL = "delete con_mezclas where cod_consolida = " & vCodigo
  glogon.Conection.Execute strSQL
  
  
  Call Bitacora("Elimina", "Consolidacion Mezcla: " & vCodigo)

  
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
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtDescConsolidacion.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDescConsolidacion_GotFocus()
 vTipoBusca = "D"
End Sub


