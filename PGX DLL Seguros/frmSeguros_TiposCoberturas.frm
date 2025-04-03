VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_TiposCoberturas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Coberturas"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6255
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   13215
      _Version        =   1441792
      _ExtentX        =   23310
      _ExtentY        =   11033
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
      SelectedItem    =   1
      Item(0).Caption =   "Coberturas"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Productos"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "vGridc"
      Item(1).Control(1)=   "Label2(0)"
      Item(1).Control(2)=   "Label2(1)"
      Item(1).Control(3)=   "cboCobertura"
      Begin FPSpreadADO.fpSpread vGridc 
         Height          =   5055
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   11535
         _Version        =   524288
         _ExtentX        =   20346
         _ExtentY        =   8916
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmSeguros_TiposCoberturas.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboCobertura 
         Height          =   330
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   6255
         _Version        =   1441792
         _ExtentX        =   11033
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5295
         Left            =   -67720
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   8055
         _Version        =   524288
         _ExtentX        =   14208
         _ExtentY        =   9340
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "frmSeguros_TiposCoberturas.frx":0709
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Productos Relacionados?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Coberturas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
   End
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   330
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   6255
      _Version        =   1441792
      _ExtentX        =   11033
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Coberturas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "frmSeguros_TiposCoberturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vRetencion As String


Private Sub cboAseguradora_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub
If cboAseguradora.ListCount = 0 Then Exit Sub

strSQL = "select codigo_Retencion from SEGUROS_ASEGURADORAS WHERE cod_aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
 vRetencion = RTrim(rs!Codigo_Retencion & "")
rs.Close

tcMain.Item(0).Selected = True
Call sbCobertura_List

End Sub


Private Sub cboCobertura_Click()
Dim strSQL As String, rs As New ADODB.Recordset


vGridc.MaxRows = 0

If vPaso Then Exit Sub
If cboAseguradora.ListCount = 0 Then Exit Sub
If cboCobertura.ListCount = 0 Then Exit Sub


strSQL = "exec spSeguros_Coberturas_Productos '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & cboCobertura.ItemData(cboCobertura.ListIndex) & "'"

vPaso = True
    Call sbCargaGrid(vGridc, 6, strSQL, True)
    vGridc.MaxRows = vGridc.MaxRows - 1
vPaso = False

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 17
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
vRetencion = ""

vPaso = True
strSQL = "select cod_aseguradora as 'IdX', rtrim(nombre) as 'ItmX' from seguros_Aseguradoras"
Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)
vPaso = False

Call cboAseguradora_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from SEGUROS_COBERTURAS " _
       & " where COD_COBERTURA = '" & vGrid.Text & "' and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "Insert SEGUROS_COBERTURAS(COD_ASEGURADORA,COD_COBERTURA,descripcion,MONTO,Activa,Registro_Usuario,Registro_Fecha)" _
         & " values('" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Cobertura:" & vGrid.Text & "_" & cboAseguradora.ItemData(cboAseguradora.ListIndex))

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update SEGUROS_COBERTURAS set descripcion = '" & vGrid.Text & "',MONTO  = "
 vGrid.Col = 3
 strSQL = strSQL & CCur(vGrid.Text) & ", Activa = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & " where COD_COBERTURA = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "' AND COD_ASEGURADORA = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Cobertura:" & vGrid.Text & "_" & cboAseguradora.ItemData(cboAseguradora.ListIndex))

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Function


Private Sub sbCobertura_List()
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass


strSQL = "select COD_COBERTURA,descripcion,MONTO,Activa" _
      & " from SEGUROS_COBERTURAS" _
      & " where cod_aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'" _
      & " order by COD_COBERTURA"
Call sbCargaGrid(vGrid, 4, strSQL)
    

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass


Select Case Item.Index

Case 0 'Coberturas
   Call sbCobertura_List
    
Case 1 'Productos Relacionados
   strSQL = "select rtrim(COD_COBERTURA) as 'IdX', rtrim(descripcion) as 'ItmX'" _
          & " from SEGUROS_COBERTURAS" _
          & " where cod_aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'" _
          & " order by COD_COBERTURA"
   vPaso = True
   Call sbCbo_Llena_New(cboCobertura, strSQL, False, True)
   vPaso = False
   
   Call cboCobertura_Click
   
End Select

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete SEGUROS_COBERTURAS" _
               & " where COD_PRODUCTO = '" & vGrid.Text & "' and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Cobertura:" & vGrid.Text & "_" & cboAseguradora.ItemData(cboAseguradora.ListIndex))
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub







Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Col = 1 Then
   vGrid.Col = 3
   If vGrid.Text = "" Then vGrid.Text = vRetencion
End If
End Sub


Private Sub vGridc_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, vAseguradora As String
Dim vProducto As String, vCobertura As String, vOpcional As Integer, vActiva As Integer

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

vCobertura = cboCobertura.ItemData(cboCobertura.ListIndex)
vAseguradora = cboAseguradora.ItemData(cboAseguradora.ListIndex)

With vGridc

     .Row = Row
     .Col = 1
     vProducto = Trim(.Text)
     
     .Col = 3
     If .Value = vbChecked Then
        vOpcional = 1
     Else
        vOpcional = 0
     End If
     
     .Col = 4
     If .Value = vbChecked Then
        vActiva = 1
     Else
        vActiva = 0
     End If
     
    strSQL = "exec spSeguros_Coberturas_Productos_Add '" & vAseguradora & "','" & vCobertura & "','" & vProducto _
            & "'," & vOpcional & "," & vActiva & ",'" & glogon.Usuario & "'"
    
    Call ConectionExecute(strSQL)
     
    Call Bitacora("Registra", "Cobertura: " & vCobertura & ", Producto: " & vProducto _
                & "(Opcional: " & vOpcional & " ¦ Activa: " & vActiva & ")")

End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


