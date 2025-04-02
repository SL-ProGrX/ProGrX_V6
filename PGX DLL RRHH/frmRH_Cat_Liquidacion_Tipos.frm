VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRH_Cat_Liquidacion_Tipos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Tipos de Liquidaciones"
   ClientHeight    =   6240
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1452
      _Version        =   1310722
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   252
      Left            =   8160
      TabIndex        =   2
      Top             =   720
      Width           =   1092
      _Version        =   1310722
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activo?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Alignment       =   1
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
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
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   9735
      _Version        =   1310722
      _ExtentX        =   17171
      _ExtentY        =   8916
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Tipos"
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "Label1(1)"
      Item(0).Control(1)=   "txtDescripcion"
      Item(0).Control(2)=   "Label1(2)"
      Item(0).Control(3)=   "cboEstadoInicial"
      Item(0).Control(4)=   "cboEstadoFinal"
      Item(0).Control(5)=   "Label1(4)"
      Item(0).Control(6)=   "chkConRespPatronal"
      Item(0).Control(7)=   "chkPreaviso"
      Item(1).Caption =   "Conceptos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "Label1(3)"
      Item(1).Control(1)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4452
         Left            =   -68200
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1310722
         _ExtentX        =   13779
         _ExtentY        =   7853
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   7452
         _Version        =   1310722
         _ExtentX        =   13144
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoInicial 
         Height          =   312
         Left            =   1800
         TabIndex        =   10
         Top             =   1080
         Width           =   2652
         _Version        =   1310722
         _ExtentX        =   4683
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoFinal 
         Height          =   312
         Left            =   1800
         TabIndex        =   11
         Top             =   1440
         Width           =   2652
         _Version        =   1310722
         _ExtentX        =   4683
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkConRespPatronal 
         Height          =   252
         Left            =   5880
         TabIndex        =   14
         Top             =   1080
         Width           =   3012
         _Version        =   1310722
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Con Responsabilidad Patronal?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkPreaviso 
         Height          =   252
         Left            =   5880
         TabIndex        =   15
         Top             =   1440
         Width           =   3012
         _Version        =   1310722
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activa Preaviso?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   1692
         _Version        =   1310722
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado Resultante"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   8
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado Inicial"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   972
         Index           =   3
         Left            =   -69880
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Seleccione los Puestos vinculados con esta Responsabilidad:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripción"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1212
      _Version        =   1310722
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo Liquidación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmRH_Cat_Liquidacion_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim vEdita  As Boolean
Dim vCodigo As String, vPaso As Boolean


Private Function fxExiste(vCodigo As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as 'Existe'" _
       & " from RH_LIQUIDACION_TIPOS where TIPO_LIQUIDACION =  '" & vCodigo & "' "
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  fxExiste = False
Else
  fxExiste = True
End If
rs.Close
End Function


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 TIPO_LIQUIDACION from RH_LIQUIDACION_TIPOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where TIPO_LIQUIDACION > '" & txtCodigo.Text & "' order by TIPO_LIQUIDACION asc"
    Else
       strSQL = strSQL & " where TIPO_LIQUIDACION < '" & txtCodigo.Text & "' order by TIPO_LIQUIDACION desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!TIPO_LIQUIDACION
      Call sbConsulta(txtCodigo.Text)
      
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 23
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 23

tcMain.Item(0).Selected = True

 
strSQL = "select ESTADO_PERSONA AS 'IdX', DESCRIPCION AS 'itmX'" _
       & " From RH_ESTADOS_TIPOS" _
       & " Where LIQUIDADO = 0 And ACTIVO = 1"
Call sbCbo_Llena_New(cboEstadoInicial, strSQL, False, True)

strSQL = "select ESTADO_PERSONA AS 'IdX', DESCRIPCION AS 'itmX'" _
       & " From RH_ESTADOS_TIPOS" _
       & " Where LIQUIDADO = 1 And ACTIVO = 1"
Call sbCbo_Llena_New(cboEstadoFinal, strSQL, False, True)


vEdita = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 950
    .Add , , "Descripción", 6750
End With

Call sbToolBarIconos(tlb, False)
Call sbToolBar(tlb, "nuevo")

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Call sbLimpiaDatos

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpiaDatos()

vCodigo = ""

tcMain.Item(0).Selected = True

txtCodigo.Text = ""
txtDescripcion.Text = ""

chkActivo.Value = xtpChecked

End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert RH_LIQUIDACION_TIPOS_CONCEPTOS(COD_CONCEPTO,TIPO_LIQUIDACION,registro_fecha,registro_usuario)" _
          & " values('" & Item.Text & "','" & vCodigo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete RH_LIQUIDACION_TIPOS_CONCEPTOS where COD_CONCEPTO = '" & Item.Text & "' and TIPO_LIQUIDACION = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
    Call sbLswLlena(txtCodigo.Text)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" And vEdita = True Then Call sbConsulta(txtCodigo.Text)
End Sub



Private Sub sbLswLlena(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem

On Error GoTo vError

vPaso = True

lsw.ListItems.Clear

strSQL = "select C.COD_CONCEPTO AS 'CODIGO', C.DESCRIPCION,ISNULL(A.COD_CONCEPTO,'') AS 'Idx'" _
       & "  from RH_CONCEPTOS C" _
       & "   LEFT JOIN RH_LIQUIDACION_TIPOS_CONCEPTOS A ON C.COD_CONCEPTO = A.COD_CONCEPTO" _
       & "   AND A.TIPO_LIQUIDACION = '" & pCodigo & "'" _
       & " WHERE C.ACTIVO = 1" _
       & " order by ISNULL(A.TIPO_LIQUIDACION,'') desc, C.COD_CONCEPTO"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      If rs!IdX <> "" Then
          itmX.Checked = vbChecked
          itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

 strSQL = "select *" _
        & " from vRH_Liquidacion_Tipos " _
        & " where TIPO_LIQUIDACION = '" & pCodigo & "'"
 Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  tcMain.Item(0).Selected = True
  vEdita = True

  txtCodigo.Text = rs!TIPO_LIQUIDACION
  vCodigo = rs!TIPO_LIQUIDACION
  
  txtDescripcion.Text = rs!Descripcion
  chkActivo.Value = rs!ACTIVO
  
  chkConRespPatronal.Value = rs!CON_RESPONSABILIDAD
  chkPreaviso.Value = rs!PREAVISO_ACTIVA

  Call sbCboAsignaDato(cboEstadoInicial, rs!ESTADO_PERSONA_DESC, True, rs!ESTADO_PERSONA)
  Call sbCboAsignaDato(cboEstadoFinal, rs!ESTADO_PERSONA_R_DESC, True, rs!ESTADO_PERSONA_RESULTANTE)
  
   
Else
  MsgBox "No se encontró registro verifique...", vbInformation
  txtCodigo.Text = ""
  txtCodigo.SetFocus
  Call sbLimpiaDatos
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        Call sbLimpiaDatos
        vEdita = False
        txtCodigo.SetFocus
       Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     Call sbToolBar(tlb, "activo")
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaDatos
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select TIPO_LIQUIDACION,descripcion from RH_LIQUIDACION_TIPOS "
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

End Select


End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset


If fxExiste(txtCodigo.Text) Then
  strSQL = "update RH_LIQUIDACION_TIPOS set descripcion = '" & Trim(txtDescripcion.Text) _
         & "', ACTIVO = " & chkActivo.Value _
         & ",  CON_RESPONSABILIDAD = " & chkConRespPatronal.Value & ", PREAVISO_ACTIVA = " & chkPreaviso.Value _
         & " , ESTADO_PERSONA = '" & cboEstadoInicial.ItemData(cboEstadoInicial.ListIndex) _
         & "', ESTADO_PERSONA_RESULTANTE = '" & cboEstadoFinal.ItemData(cboEstadoFinal.ListIndex) _
         & "' where TIPO_LIQUIDACION = '" & vCodigo & "' "
         
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Tipo de Liquidacion Laboral: " & vCodigo)

Else
  vCodigo = txtCodigo.Text

   strSQL = "insert into RH_LIQUIDACION_TIPOS(TIPO_LIQUIDACION,descripcion,ACTIVO,ESTADO_PERSONA" _
          & ",ESTADO_PERSONA_RESULTANTE, CON_RESPONSABILIDAD, PREAVISO_ACTIVA,REGISTRO_USUARIO, REGISTRO_FECHA)" _
          & " values('" & vCodigo & "','" & Trim(txtDescripcion.Text) & "', " & chkActivo.Value & ",'" _
          & cboEstadoInicial.ItemData(cboEstadoInicial.ListIndex) & "','" & cboEstadoFinal.ItemData(cboEstadoFinal.ListIndex) _
          & "'," & chkConRespPatronal.Value & "," & chkPreaviso.Value & ",'" & glogon.Usuario & " ',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Tipo de Liquidacion Laboral: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida()

fxValida = True

If Trim(txtCodigo) = "" Then fxValida = False
If Trim(txtDescripcion) = "" Then fxValida = False

End Function



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tcMain.Item(0).Selected = True
    txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Liq Tipo"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "TIPO_LIQUIDACION"
   gBusquedas.Orden = "TIPO_LIQUIDACION"
   gBusquedas.Consulta = "select TIPO_LIQUIDACION,descripcion from RH_LIQUIDACION_TIPOS"
   frmBusquedas.Show vbModal
   txtCodigo.Text = gBusquedas.Resultado
   
   tcMain.Item(0).Selected = True
   txtDescripcion.SetFocus
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstadoInicial.SetFocus
End Sub



