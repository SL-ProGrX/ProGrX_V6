VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCR_ComitesAutorizadores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizadores de Comites"
   ClientHeight    =   7020
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11916
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   11916
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5652
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   11772
      _Version        =   1245187
      _ExtentX        =   20764
      _ExtentY        =   9970
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
      ItemCount       =   3
      Item(0).Caption =   "Puestos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGridPuestos"
      Item(1).Caption =   "Personas"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Item(2).Caption =   "Asignación"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "cboComites"
      Item(2).Control(1)=   "Label2(6)"
      Item(2).Control(2)=   "tcAux"
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   4932
         Left            =   -67720
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   8292
         _Version        =   1245187
         _ExtentX        =   14626
         _ExtentY        =   8700
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
         Item(0).Caption =   "Miembros"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "lswMiembros"
         Item(1).Caption =   "Autorizadores"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswAutorizadores"
         Begin XtremeSuiteControls.ListView lswAutorizadores 
            Height          =   4452
            Left            =   -69880
            TabIndex        =   7
            Top             =   360
            Visible         =   0   'False
            Width           =   8172
            _Version        =   1245187
            _ExtentX        =   14414
            _ExtentY        =   7853
            _StockProps     =   77
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
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ListView lswMiembros 
            Height          =   4452
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   8172
            _Version        =   1245187
            _ExtentX        =   14414
            _ExtentY        =   7853
            _StockProps     =   77
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
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   16
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5100
         Left            =   -69880
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   11652
         _Version        =   524288
         _ExtentX        =   20553
         _ExtentY        =   8996
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
         MaxCols         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_ComitesAutorizadores.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridPuestos 
         Height          =   5100
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   7212
         _Version        =   524288
         _ExtentX        =   12721
         _ExtentY        =   8996
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
         MaxCols         =   493
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_ComitesAutorizadores.frx":0639
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboComites 
         Height          =   312
         Left            =   -67600
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   8172
         _Version        =   1245187
         _ExtentX        =   14415
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comites"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   6
         Left            =   -68920
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizadores del Comité de Resolución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   1920
      TabIndex        =   0
      Top             =   300
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmCR_ComitesAutorizadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mUlltimoPuestoSel As String, strUltimaSelTipo As String, mListaPuestos As String
Dim vPaso  As Boolean


Private Sub Form_Activate()
    vModulo = 3
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rs As New ADODB.Recordset
    
    vModulo = 3
    
    tcMain.Item(0).Selected = True
    
    
    vGrid.AppearanceStyle = fxGridStyle
    vGridPuestos.AppearanceStyle = vGrid.AppearanceStyle
    
    Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


    With lswAutorizadores.ColumnHeaders
         .Clear
         .Add , , "Identificación", 2200
         .Add , , "Nombre", 5200
    End With
    
    
    Call Formularios(Me)
    Call RefrescaTags(Me)
    
  On Error GoTo vError
  
    strSQL = "select rtrim(ID_PUESTO) + ' - ' + descripcion as Puesto" _
           & " from CRD_COMITES_MIEMBROS_PUESTOS" _
           & " order by ID_PUESTO"
    
    Call OpenRecordSet(rs, strSQL)
        
        If Not rs.EOF And mUlltimoPuestoSel = "" Then
         mUlltimoPuestoSel = rs!PUESTO
        End If
        
        mListaPuestos = ""
        Do While Not rs.EOF
            If Len(mListaPuestos) = 0 Then
              mListaPuestos = rs!PUESTO
            Else
              mListaPuestos = mListaPuestos & Chr$(9) & rs!PUESTO
            End If
          rs.MoveNext
        Loop
    rs.Close
    
    Call sbPuestos_Load
    
 Exit Sub
vError:
 
End Sub

Private Sub sbPuestos_Load()
Dim strSQL As String

    tcMain.Item(0).Selected = True
    
    Me.MousePointer = vbHourglass


    strSQL = "select ID_PUESTO,DESCRIPCION from CRD_COMITES_MIEMBROS_PUESTOS order by ID_PUESTO"
    Call sbCargaGrid(vGridPuestos, 2, strSQL)
    
    
    Me.MousePointer = vbDefault

End Sub

Private Sub lswAutorizadores_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError
    
    If Item.Checked Then
      strSQL = "insert CRD_COMITES_AUTORIZADORES(CEDULA,ID_COMITE,REGISTRO_FECHA, REGISTRO_USUARIO) values('" & Item.Text _
             & "'," & cboComites.ItemData(cboComites.ListIndex) & ", dbo.Mygetdate(), '" & glogon.Usuario & "')"
    Else
      strSQL = "delete CRD_COMITES_AUTORIZADORES where CEDULA = '" & Item.Text _
             & "' and ID_COMITE = '" & cboComites.ItemData(cboComites.ListIndex) & "'"
    End If
    Call ConectionExecute(strSQL)
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxGuardarPuesto() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardarPuesto = 0
vGridPuestos.Row = vGridPuestos.ActiveRow
vGridPuestos.col = 1

strSQL = "select isnull(count(*),0) as Existe from CRD_COMITES_MIEMBROS_PUESTOS" _
       & " where ID_PUESTO = '" & vGridPuestos.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGridPuestos.Text) = "" Then Exit Function
  
  strSQL = "insert into CRD_COMITES_MIEMBROS_PUESTOS(ID_PUESTO,DESCRIPCION) values('" _
         & UCase(vGridPuestos.Text) & "','"
  vGridPuestos.col = 2
  strSQL = strSQL & UCase(vGridPuestos.Text) & "')"

  Call ConectionExecute(strSQL)

  vGridPuestos.col = 1
  Call Bitacora("Registra", "Comites Puestos Miembros Autorizadores: " & vGridPuestos.Text)

Else 'Actualizar

 vGridPuestos.col = 2
 strSQL = "update CRD_COMITES_MIEMBROS_PUESTOS set descripcion = '" & vGridPuestos.Text & "'"
 strSQL = strSQL & " where ID_PUESTO = '"
 vGridPuestos.col = 1
 strSQL = strSQL & vGridPuestos.Text & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Comites Puestos Miembros Autorizadores: " & vGridPuestos.Text)


End If
rs.Close

fxGuardarPuesto = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function
Private Sub sbPersonas_Load()
Dim strSQL As String

On Error GoTo vError

    strSQL = "select M.CEDULA,M.NOMBRE,M.USUARIO,isnull(rtrim(P.ID_PUESTO) + ' - ' + P.DESCRIPCION,'') as PUESTO, " _
          & "  CASE M.ESTADO WHEN 'A' THEN 1 ELSE 0 END AS ESTADO, FECHA_ACTIVA, USUARIO_ACTIVA, FECHA_BLOQUEO, USUARIO_BLOQUEO" _
          & "  from CRD_COMITES_MIEMBROS M inner join CRD_COMITES_MIEMBROS_PUESTOS P on M.ID_PUESTO = P.ID_PUESTO" _
          & " order by M.ID_PUESTO"
          
    Call sbCargaGridLocal(vGrid, 5, strSQL)
    
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, pMov As String

If vPaso Then Exit Sub

On Error GoTo vError
    
    If Item.Checked Then
      pMov = "E"
    Else
      pMov = "S"
    End If
    
    strSQL = "exec spCrd_Comites_Miembros_Add " & cboComites.ItemData(cboComites.ListIndex) _
            & ",'" & Item.Text & "','" & glogon.Usuario & "','" & pMov & "'"
    Call ConectionExecute(strSQL)
    
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
  Call sbAutorizadores_Load
End If

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Puestos
    Call sbPuestos_Load
  Case 1 'Autorizadores
    Call sbPersonas_Load
  Case 2 'Asignación
    Call sbComites_Cbo_Load
End Select

End Sub

Private Sub vGridPuestos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGridPuestos.ActiveCol = vGridPuestos.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardarPuesto
  If i = 0 Then Exit Sub
  vGridPuestos.Row = vGridPuestos.ActiveRow
  If vGridPuestos.MaxRows <= vGridPuestos.ActiveRow Then
    vGridPuestos.MaxRows = vGridPuestos.MaxRows + 1
    vGridPuestos.Row = vGridPuestos.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridPuestos.MaxRows = vGridPuestos.MaxRows + 1
    vGridPuestos.InsertRows vGridPuestos.ActiveRow, 1
    vGridPuestos.Row = vGridPuestos.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este puesto", vbYesNo)
     If i = vbYes Then
        vGridPuestos.Row = vGridPuestos.ActiveRow
        vGridPuestos.col = 1
        strSQL = "delete CRD_COMITES_MIEMBROS_PUESTOS where ID_PUESTO = '" & vGridPuestos.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGridPuestos.Text
        vGridPuestos.col = 1

        vGridPuestos.DeleteRows vGridPuestos.ActiveRow, 1
        vGridPuestos.MaxRows = vGridPuestos.MaxRows - 1

        If vGridPuestos.MaxRows <= 0 Then
          vGridPuestos.MaxRows = 1
        End If

     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim strResTipo As String, vNota As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 4
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mListaPuestos
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mUlltimoPuestoSel
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1 'Cedula
       vGrid.Text = CStr(rs!Cedula)
     Case 2 'Nombre
       vGrid.Text = CStr(rs!Nombre)
     Case 3 'USUARIO
       vGrid.Text = CStr(rs!Usuario)
       
       vGrid.TextTip = TextTipFixed
       vGrid.TextTipDelay = 1000
                
       vGrid.CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
       vGrid.CellNoteIndicatorColor = vbRed
                  
       vGrid.CellNote = "Activación: " & Format(rs!FECHA_ACTIVA, "dd/mm/yyyy") & " [" & rs!USUARIO_ACTIVA & "]" & vbCrLf & _
                "Bloqueo: " & Format(rs!FECHA_BLOQUEO, "dd/mm/yyyy") & " [" & rs!USUARIO_BLOQUEO & "]"
                     
     Case 4 '
        vGrid.Text = rs!PUESTO
     Case 5 'Activo
       vGrid.Text = CStr(rs!Estado)
     Case Else
    End Select
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 4
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mListaPuestos
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mUlltimoPuestoSel

Me.MousePointer = vbDefault

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from CRD_COMITES_MIEMBROS " _
       & " where CEDULA = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into CRD_COMITES_MIEMBROS(CEDULA,NOMBRE,USUARIO,ID_PUESTO,ESTADO,FECHA_ACTIVA,USUARIO_ACTIVA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 4
  strSQL = strSQL & "'" & SIFGlobal.fxCodText(vGrid.Text) & "','A',dbo.MyGetdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Autorizador de comites : " & vGrid.Text)
  
  vGrid.col = 5
  vGrid.Value = 1

Else 'Actualizar

   vGrid.col = 1
   rs.Close
   strSQL = "select Estado from CRD_COMITES_MIEMBROS " _
           & " where CEDULA = '" & vGrid.Text & "'"
   Call OpenRecordSet(rs, strSQL)

  vGrid.col = 2
  strSQL = "update CRD_COMITES_MIEMBROS set NOMBRE = '" & vGrid.Text
  
  vGrid.col = 3
  strSQL = strSQL & "',USUARIO='" & vGrid.Text
  vGrid.col = 4
  strSQL = strSQL & "',ID_PUESTO='" & SIFGlobal.fxCodText(vGrid.Text) & "'"
  vGrid.col = 5
  
  If rs!Estado = "A" And vGrid.Value = 0 Then
          strSQL = strSQL & ",ESTADO='B',USUARIO_BLOQUEO = '" & glogon.Usuario & "',FECHA_BLOQUEO = dbo.MyGetdate()"
  Else
    If rs!Estado = "B" And vGrid.Value = 1 Then
        strSQL = strSQL & ",ESTADO='A',USUARIO_ACTIVA = '" & glogon.Usuario & "',FECHA_ACTIVA = dbo.MyGetdate()"
    End If
  End If
  
  strSQL = strSQL & " where cedula = '"
  vGrid.col = 1
  strSQL = strSQL & vGrid.Text & "'"
  Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Autorizador de comites : " & vGrid.Text)

End If
rs.Close

vGrid.col = 4
mUlltimoPuestoSel = vGrid.Text


fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

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
  
    vGrid.col = 4
    vGrid.CellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mListaPuestos
    vGrid.TypeComboBoxEditable = False
    vGrid.Text = mUlltimoPuestoSel
  
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow

    vGrid.col = 4
    vGrid.CellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mListaPuestos
    vGrid.TypeComboBoxEditable = False
    vGrid.Text = mUlltimoPuestoSel

End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este autorizador", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete CRD_COMITES_MIEMBROS where CEDULA = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Autorizador de comites : " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1

        If vGrid.MaxRows <= 0 Then
          vGrid.MaxRows = 1
        End If

     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbComites_Cbo_Load()
Dim strSQL As String
    
vPaso = True
    strSQL = "Select id_comite as 'IdX',rtrim(descripcion) as 'ItmX' from comites"
    Call sbCbo_Llena_New(cboComites, strSQL, False, True)
vPaso = False

Call cboComites_Click

End Sub


Private Sub sbAutorizadores_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub
If cboComites.ListCount <= 0 Then Exit Sub

On Error GoTo vError

tcAux.Item(1).Selected = True

vPaso = True

With lswAutorizadores
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Identificación", 2100
    .ColumnHeaders.Add , , "Nombre", 4100
    
    strSQL = "exec spCrd_Comites_Miembros_Autoriza_Consulta " & cboComites.ItemData(cboComites.ListIndex)
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     Set itmX = .ListItems.Add(, , rs!Cedula)
         itmX.SubItems(1) = rs!Nombre
         If rs!Asignado = 1 Then
            itmX.Checked = vbChecked
            itmX.ForeColor = vbBlue
         End If
     rs.MoveNext
    Loop

End With

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboComites_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub
If cboComites.ListCount <= 0 Then Exit Sub

On Error GoTo vError

tcAux.Item(0).Selected = True

vPaso = True

With lswMiembros
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Identificación", 2100
    .ColumnHeaders.Add , , "Nombre", 4100
    
    strSQL = "exec spCrd_Comites_Miembros_Consulta " & cboComites.ItemData(cboComites.ListIndex)
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     Set itmX = .ListItems.Add(, , rs!Cedula)
         itmX.SubItems(1) = rs!Nombre
         If rs!Asignado = 1 Then
            itmX.Checked = vbChecked
            itmX.ForeColor = vbBlue
         End If
     rs.MoveNext
    Loop

End With

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
