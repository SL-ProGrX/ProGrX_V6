VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmPreaClasificaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clasificaciones : Integrador"
   ClientHeight    =   6480
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   8772
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8772
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5172
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8532
      _Version        =   1245187
      _ExtentX        =   15049
      _ExtentY        =   9123
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
      ItemCount       =   6
      SelectedItem    =   1
      Item(0).Caption =   "Razones"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Garantía"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "cboGarantia"
      Item(1).Control(1)=   "lswGarantia"
      Item(1).Control(2)=   "vGridGarantia"
      Item(1).Control(3)=   "scTitulo"
      Item(1).Control(4)=   "btnRefresh"
      Item(2).Caption =   "Morosidad"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGridMora"
      Item(3).Caption =   "Capacidad"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "vGridCapacidad"
      Item(4).Caption =   "Endeudamiento"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "vGridEndeudamiento"
      Item(5).Caption =   "Historial de Pago"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "vGridHistorial"
      Begin XtremeSuiteControls.PushButton btnRefresh 
         Height          =   312
         Left            =   7800
         TabIndex        =   11
         Top             =   2436
         Width           =   432
         _Version        =   1245187
         _ExtentX        =   762
         _ExtentY        =   550
         _StockProps     =   79
         Appearance      =   6
         Picture         =   "frmPreaClasificaciones.frx":0000
      End
      Begin XtremeSuiteControls.ListView lswGarantia 
         Height          =   2172
         Left            =   2760
         TabIndex        =   8
         Top             =   2880
         Width           =   4932
         _Version        =   1245187
         _ExtentX        =   8700
         _ExtentY        =   3831
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
         ShowBorder      =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4572
         Left            =   -69280
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   6972
         _Version        =   524288
         _ExtentX        =   12298
         _ExtentY        =   8064
         _StockProps     =   64
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmPreaClasificaciones.frx":0700
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridGarantia 
         Height          =   1692
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   8052
         _Version        =   524288
         _ExtentX        =   14203
         _ExtentY        =   2985
         _StockProps     =   64
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmPreaClasificaciones.frx":0C4E
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridMora 
         Height          =   4452
         Left            =   -69760
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   7932
         _Version        =   524288
         _ExtentX        =   13991
         _ExtentY        =   7853
         _StockProps     =   64
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmPreaClasificaciones.frx":1188
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridCapacidad 
         Height          =   4572
         Left            =   -68920
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   6372
         _Version        =   524288
         _ExtentX        =   11239
         _ExtentY        =   8064
         _StockProps     =   64
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmPreaClasificaciones.frx":182F
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridEndeudamiento 
         Height          =   4572
         Left            =   -68920
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   6372
         _Version        =   524288
         _ExtentX        =   11239
         _ExtentY        =   8064
         _StockProps     =   64
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmPreaClasificaciones.frx":1E30
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridHistorial 
         Height          =   4572
         Left            =   -69640
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   7932
         _Version        =   524288
         _ExtentX        =   13991
         _ExtentY        =   8064
         _StockProps     =   64
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmPreaClasificaciones.frx":2431
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboGarantia 
         Height          =   312
         Left            =   2760
         TabIndex        =   9
         Top             =   2436
         Width           =   4932
         _Version        =   1245187
         _ExtentX        =   8700
         _ExtentY        =   550
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   8052
         _Version        =   1245187
         _ExtentX        =   14203
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Asigna Razón de Garantía"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clasificaciones Crediticias de la Persona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPreaClasificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRazones As String, vPaso As Boolean

Private Sub btnRefresh_Click()
Call sbGarantia_Cbo_Load
End Sub

Private Sub cboGarantia_Click()

If vPaso Then Exit Sub

Call sbGarantia_LoadLista(cboGarantia.ItemData(cboGarantia.ListIndex))

End Sub

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_Load()

vModulo = 3 'Modulo de Credito

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswGarantia.ColumnHeaders
    .Clear
    .Add , , "Garantía", lswGarantia.Width - 250
End With

Call Formularios(Me)
Call RefrescaTags(Me)


tcMain.Item(0).Selected = True
Call sbRazones_Load

End Sub


Private Sub sbRazones_Cbos_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

mRazones = ""

strSQL = "select rtrim(cod_razon) + ' - ' + rtrim(descripcion) as iTemX from Crd_Clasificacion_Razon" _
      & " order by cod_razon"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  mRazones = mRazones & Chr$(9) & rs!iTemX
  rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

 
End Sub

Private Sub sbRazones_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "select cod_razon,descripcion,color from Crd_Clasificacion_Razon" _
      & " order by cod_razon"
Call sbCargaGrid(vGrid, 3, strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbGarantia_Cbo_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "select cod_garantia as 'IdX', descripcion as 'ItmX'" _
       & " from Crd_Clasificacion_Garantia" _
       & " order by cod_Garantia"

vPaso = True

Call sbCbo_Llena_New(cboGarantia, strSQL, False, True)

vPaso = False

Call cboGarantia_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGarantia_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "select A.cod_garantia,A.descripcion, rtrim(B.cod_Razon) + ' - ' + rtrim(B.descripcion)" _
       & " from Crd_Clasificacion_Garantia A inner join Crd_Clasificacion_Razon B on A.cod_Razon = B.Cod_Razon" _
      & " order by A.cod_Garantia"
Call sbCargaGridLocal(vGridGarantia, 3, strSQL)

Call sbGarantia_Cbo_Load

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbMora_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "select A.cod_mora, case " _
       & " when A.tipo = 'A' then 'Al Día'" _
       & " when A.tipo = 'M' then 'Mora'" _
       & " when A.tipo = 'C' then 'Cobro (Ejecutado)'" _
       & " when A.tipo = 'I' then 'Incobrable' end as Tipo " _
       & ",A.desde,A.hasta,rtrim(B.cod_Razon) + ' - ' + rtrim(B.descripcion)" _
       & " from Cbr_Clasificacion_Mora A inner join Crd_Clasificacion_Razon B on A.cod_Razon = B.Cod_Razon" _
       & " order by A.cod_mora"
Call sbCargaGridLocal(vGridMora, 5, strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCapacidad_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "select A.cod_capacidad,A.desde,A.hasta,rtrim(B.cod_Razon) + ' - ' + rtrim(B.descripcion)" _
       & " from Crd_Clasificacion_Capacidad A inner join Crd_Clasificacion_Razon B on A.cod_Razon = B.Cod_Razon" _
       & " order by A.cod_capacidad"
Call sbCargaGridLocal(vGridCapacidad, 4, strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbEndeudamiento_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "select A.cod_endeudamiento,A.desde,A.hasta,rtrim(B.cod_Razon) + ' - ' + rtrim(B.descripcion)" _
       & " from Crd_Clasificacion_endeudamiento A inner join Crd_Clasificacion_Razon B on A.cod_Razon = B.Cod_Razon" _
       & " order by A.cod_endeudamiento"
Call sbCargaGridLocal(vGridEndeudamiento, 4, strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbHistorial_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "select A.cod_historial,A.descripcion, rtrim(B.cod_Razon) + ' - ' + rtrim(B.descripcion)" _
       & " from Crd_Clasificacion_historial A inner join Crd_Clasificacion_Razon B on A.cod_Razon = B.Cod_Razon" _
      & " order by A.cod_historial"
Call sbCargaGridLocal(vGridHistorial, 3, strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbCargaGridLocal(vGridTmp As Object, vGridTmpMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

vGridTmp.MaxCols = vGridTmpMaxCol
vGridTmp.MaxRows = 1
vGridTmp.Row = vGridTmp.MaxRows
For i = 1 To vGridTmp.MaxCols
 vGridTmp.Col = i
 vGridTmp.Text = ""
Next i

  vGridTmp.Col = vGridTmp.MaxCols
  vGridTmp.CellType = CellTypeComboBox
  vGridTmp.TypeComboBoxList = mRazones
  vGridTmp.TypeComboBoxEditable = False


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGridTmp.Row = vGridTmp.MaxRows
  
  vGridTmp.Col = vGridTmp.MaxCols
  vGridTmp.CellType = CellTypeComboBox
  vGridTmp.TypeComboBoxList = mRazones
  vGridTmp.TypeComboBoxEditable = False
  
  For i = 1 To vGridTmp.MaxCols
    vGridTmp.Col = i
    vGridTmp.Text = CStr(rs.Fields(i - 1).Value)
  Next i
  vGridTmp.MaxRows = vGridTmp.MaxRows + 1
  
  rs.MoveNext
Loop
    
  '' 05/04/2011 No cargaba el combo en la última línea
  ''******
  vGridTmp.Row = vGridTmp.MaxRows
  vGridTmp.Col = vGridTmp.MaxCols
  vGridTmp.CellType = CellTypeComboBox
  vGridTmp.TypeComboBoxList = mRazones
  vGridTmp.TypeComboBoxEditable = False
  ''******
  
rs.Close
End Sub


Private Function fxGarantiaCheck(xCodGarantia As String, xGarantia As String) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select * from crd_clasificacion_garantia_dt" _
       & " where cod_garantia = '" & xCodGarantia & "' and Garantia = '" & xGarantia & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
  fxGarantiaCheck = False
Else
  fxGarantiaCheck = True
End If
rsX.Close

End Function

Private Sub sbGarantia_LoadLista(pGarantiaRazon As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

vPaso = True

lswGarantia.ListItems.Clear
 
strSQL = "select Gt.GARANTIA, Gt.DESCRIPCION, case when isnull(Gr.COD_GARANTIA,'') = '' then 0 else 1 end 'Checked'" _
       & " from  CRD_GARANTIA_TIPOS Gt" _
       & "        left join CRD_CLASIFICACION_GARANTIA_DT Gr on Gt.GARANTIA = Gr.GARANTIA and Gr.COD_GARANTIA = '" & pGarantiaRazon & "'" _
       & " order by Gr.COD_GARANTIA desc, Gt.DESCRIPCION"
 
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lswGarantia.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!GARANTIA
     itmX.Checked = IIf((rs!Checked = 1), True, False)
 rs.MoveNext
Loop
rs.Close


vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxRazones_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxRazones_Guardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Crd_Clasificacion_Razon " _
       & " where cod_razon = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Clasificacion_Razon(cod_razon,descripcion,color) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "PreAnalisis (Razon) : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Crd_Clasificacion_Razon set descripcion = '" & vGrid.Text & "',color = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "' where cod_razon = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "PreAnalisis (Razon) : " & vGrid.Text)

End If
rs.Close

fxRazones_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Function fxGarantia_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGarantia_Guardar = 0
vGridGarantia.Row = vGridGarantia.ActiveRow
vGridGarantia.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Crd_Clasificacion_Garantia " _
       & " where cod_Garantia = '" & vGridGarantia.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGridGarantia.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Clasificacion_Garantia(cod_Garantia,descripcion,cod_razon) values('" _
         & vGridGarantia.Text & "','"
  vGridGarantia.Col = 2
  strSQL = strSQL & vGridGarantia.Text & "','"
  vGridGarantia.Col = 3
  strSQL = strSQL & SIFGlobal.fxCodText(vGridGarantia.Text) & "')"

  Call ConectionExecute(strSQL)

  vGridGarantia.Col = 1
  Call Bitacora("Registra", "Clasificacion Garantía : " & vGridGarantia.Text)

Else 'Actualizar

 vGridGarantia.Col = 2
 strSQL = "update Crd_Clasificacion_Garantia set descripcion = '" & vGridGarantia.Text & "',cod_razon = '"
 vGridGarantia.Col = 3
 strSQL = strSQL & SIFGlobal.fxCodText(vGridGarantia.Text) & "' where cod_Garantia = '"
 vGridGarantia.Col = 1
 strSQL = strSQL & vGridGarantia.Text & "'"
 Call ConectionExecute(strSQL)

 vGridGarantia.Col = 1
 Call Bitacora("Modifica", "Clasificacion Garantía : " & vGridGarantia.Text)

End If
rs.Close

fxGarantia_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Function fxMora_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxMora_Guardar = 0
vGridMora.Row = vGridMora.ActiveRow
vGridMora.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Cbr_Clasificacion_Mora " _
       & " where cod_mora = '" & vGridMora.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGridMora.Text) = "" Then Exit Function
  
  strSQL = "insert into Cbr_Clasificacion_Mora(cod_mora,tipo,desde,hasta,cod_razon) values('" _
         & vGridMora.Text & "','"
  vGridMora.Col = 2
  strSQL = strSQL & Mid(vGridMora.Text, 1, 1) & "',"
  vGridMora.Col = 3
  strSQL = strSQL & CCur(vGridMora.Text) & ","
  vGridMora.Col = 4
  strSQL = strSQL & CCur(vGridMora.Text) & ",'"
  vGridMora.Col = 5
  strSQL = strSQL & SIFGlobal.fxCodText(vGridMora.Text) & "')"

  Call ConectionExecute(strSQL)

  vGridMora.Col = 1
  Call Bitacora("Registra", "Clasificacion Mora : " & vGridMora.Text)

Else 'Actualizar

 vGridMora.Col = 2
 strSQL = "update Cbr_Clasificacion_Mora set tipo = '" & Mid(vGridMora.Text, 1, 1) _
        & "',desde = "
 vGridMora.Col = 3
 strSQL = strSQL & CCur(vGridMora.Text) & ",hasta = "
 vGridMora.Col = 4
 strSQL = strSQL & CCur(vGridMora.Text) & ",cod_razon = '"
 vGridMora.Col = 5
 strSQL = strSQL & SIFGlobal.fxCodText(vGridMora.Text) & "' where cod_mora = '"
 vGridMora.Col = 1
 strSQL = strSQL & vGridMora.Text & "'"
 Call ConectionExecute(strSQL)

 vGridMora.Col = 1
 Call Bitacora("Modifica", "Clasificacion Mora : " & vGridMora.Text)

End If
rs.Close

fxMora_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Function fxCapacidad_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxCapacidad_Guardar = 0
vGridCapacidad.Row = vGridCapacidad.ActiveRow
vGridCapacidad.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Crd_Clasificacion_Capacidad " _
       & " where cod_capacidad = '" & vGridCapacidad.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGridCapacidad.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Clasificacion_Capacidad(cod_capacidad,desde,hasta,cod_razon) values('" _
         & vGridCapacidad.Text & "',"
  vGridCapacidad.Col = 2
  strSQL = strSQL & CCur(vGridCapacidad.Text) & ","
  vGridCapacidad.Col = 3
  strSQL = strSQL & CCur(vGridCapacidad.Text) & ",'"
  vGridCapacidad.Col = 4
  strSQL = strSQL & SIFGlobal.fxCodText(vGridCapacidad.Text) & "')"

  Call ConectionExecute(strSQL)

  vGridCapacidad.Col = 1
  Call Bitacora("Registra", "Clasificacion Capacidad : " & vGridCapacidad.Text)

Else 'Actualizar

 strSQL = "update Crd_Clasificacion_Capacidad set desde = "
 vGridCapacidad.Col = 2
 strSQL = strSQL & CCur(vGridCapacidad.Text) & ",hasta = "
 vGridCapacidad.Col = 3
 strSQL = strSQL & CCur(vGridCapacidad.Text) & ",cod_razon = '"
 vGridCapacidad.Col = 4
 strSQL = strSQL & SIFGlobal.fxCodText(vGridCapacidad.Text) & "' where cod_capacidad = '"
 vGridCapacidad.Col = 1
 strSQL = strSQL & vGridCapacidad.Text & "'"
 Call ConectionExecute(strSQL)

 vGridCapacidad.Col = 1
 Call Bitacora("Modifica", "Clasificacion Capacidad : " & vGridCapacidad.Text)

End If
rs.Close

fxCapacidad_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Function fxEndeudamiento_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

 

fxEndeudamiento_Guardar = 0
vGridEndeudamiento.Row = vGridEndeudamiento.ActiveRow
vGridEndeudamiento.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Crd_Clasificacion_Endeudamiento " _
       & " where cod_Endeudamiento  = '" & vGridEndeudamiento.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGridEndeudamiento.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Clasificacion_Endeudamiento (cod_Endeudamiento ,desde,hasta,cod_razon) values('" _
         & vGridEndeudamiento.Text & "',"
  vGridEndeudamiento.Col = 2
  strSQL = strSQL & CCur(vGridEndeudamiento.Text) & ","
  vGridEndeudamiento.Col = 3
  strSQL = strSQL & CCur(vGridEndeudamiento.Text) & ",'"
  vGridEndeudamiento.Col = 4
  strSQL = strSQL & SIFGlobal.fxCodText(vGridEndeudamiento.Text) & "')"

  Call ConectionExecute(strSQL)

  vGridEndeudamiento.Col = 1
  Call Bitacora("Registra", "Clasificacion Endeudamiento  : " & vGridEndeudamiento.Text)

Else 'Actualizar

 strSQL = "update Crd_Clasificacion_Endeudamiento  set desde = "
 vGridEndeudamiento.Col = 2
 strSQL = strSQL & CCur(vGridEndeudamiento.Text) & ",hasta = "
 vGridEndeudamiento.Col = 3
 strSQL = strSQL & CCur(vGridEndeudamiento.Text) & ",cod_razon = '"
 vGridEndeudamiento.Col = 4
 strSQL = strSQL & SIFGlobal.fxCodText(vGridEndeudamiento.Text) & "' where cod_Endeudamiento  = '"
 vGridEndeudamiento.Col = 1
 strSQL = strSQL & vGridEndeudamiento.Text & "'"
 Call ConectionExecute(strSQL)

 vGridEndeudamiento.Col = 1
 Call Bitacora("Modifica", "Clasificacion Endeudamiento  : " & vGridEndeudamiento.Text)

End If
rs.Close

fxEndeudamiento_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Function fxHistorial_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxHistorial_Guardar = 0
vGridHistorial.Row = vGridHistorial.ActiveRow
vGridHistorial.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Crd_Clasificacion_Historial " _
       & " where cod_historial = '" & vGridHistorial.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGridHistorial.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Clasificacion_Historial(cod_historial,descripcion,cod_razon) values('" _
         & vGridHistorial.Text & "','"
  vGridHistorial.Col = 2
  strSQL = strSQL & vGridHistorial.Text & "','"
  vGridHistorial.Col = 3
  strSQL = strSQL & SIFGlobal.fxCodText(vGridHistorial.Text) & "')"

  Call ConectionExecute(strSQL)

  vGridHistorial.Col = 1
  Call Bitacora("Registra", "Clasificacion Historial : " & vGridHistorial.Text)

Else 'Actualizar

 vGridHistorial.Col = 2
 strSQL = "update Crd_Clasificacion_Historial set descripcion = '" & vGridHistorial.Text & "',cod_razon = '"
 vGridHistorial.Col = 3
 strSQL = strSQL & SIFGlobal.fxCodText(vGridHistorial.Text) & "' where cod_historial = '"
 vGridHistorial.Col = 1
 strSQL = strSQL & vGridHistorial.Text & "'"
 Call ConectionExecute(strSQL)

 vGridHistorial.Col = 1
 Call Bitacora("Modifica", "Clasificacion Historial : " & vGridHistorial.Text)

End If
rs.Close

fxHistorial_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub sbRazones_Borrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   strSQL = "delete Crd_Clasificacion_Razon where cod_razon = '" & vGrid.Text & "'"
   Call ConectionExecute(strSQL)
   
   vGrid.Col = 1
   Call Bitacora("Elimina", "PreAnalisis (Razon) : " & vGrid.Text)
   
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
   If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGarantia_Borrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGridGarantia.Row = vGridGarantia.ActiveRow
   vGridGarantia.Col = 1
   
   'Eliminar detalle Primero
   
   strSQL = "delete Crd_Clasificacion_Garantia_Dt where cod_garantia = '" & vGridGarantia.Text & "'"
   Call ConectionExecute(strSQL)
   
   strSQL = "delete Crd_Clasificacion_Garantia where cod_garantia = '" & vGridGarantia.Text & "'"
   Call ConectionExecute(strSQL)
   
   vGridGarantia.Col = 1
   Call Bitacora("Elimina", "Clasificacion Garantia : " & vGridGarantia.Text)
   
   vGridGarantia.DeleteRows vGridGarantia.ActiveRow, 1
   vGridGarantia.MaxRows = vGridGarantia.MaxRows - 1
   If vGridGarantia.MaxRows = 0 Then vGridGarantia.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbMora_Borrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGridMora.Row = vGridMora.ActiveRow
   vGridMora.Col = 1
   
   strSQL = "delete Cbr_Clasificacion_Mora where cod_Mora = '" & vGridMora.Text & "'"
   Call ConectionExecute(strSQL)
   
   vGridMora.Col = 1
   Call Bitacora("Elimina", "Clasificacion Mora : " & vGridMora.Text)
   
   vGridMora.DeleteRows vGridMora.ActiveRow, 1
   vGridMora.MaxRows = vGridMora.MaxRows - 1
   If vGridMora.MaxRows = 0 Then vGridMora.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCapacidad_Borrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGridCapacidad.Row = vGridCapacidad.ActiveRow
   vGridCapacidad.Col = 1
   
   strSQL = "delete Crd_Clasificacion_Capacidad where cod_capacidad = '" & vGridCapacidad.Text & "'"
   Call ConectionExecute(strSQL)
   
   vGridCapacidad.Col = 1
   Call Bitacora("Elimina", "Clasificacion Capacidad : " & vGridCapacidad.Text)
   
   vGridCapacidad.DeleteRows vGridCapacidad.ActiveRow, 1
   vGridCapacidad.MaxRows = vGridCapacidad.MaxRows - 1
   If vGridCapacidad.MaxRows = 0 Then vGridCapacidad.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbEndeudamiento_Borrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGridEndeudamiento.Row = vGridEndeudamiento.ActiveRow
   vGridEndeudamiento.Col = 1
   
   strSQL = "delete Crd_Clasificacion_Endeudamiento where cod_Endeudamiento = '" & vGridEndeudamiento.Text & "'"
   Call ConectionExecute(strSQL)
   
   vGridEndeudamiento.Col = 1
   Call Bitacora("Elimina", "Clasificacion Endeudamiento : " & vGridEndeudamiento.Text)
   
   vGridEndeudamiento.DeleteRows vGridEndeudamiento.ActiveRow, 1
   vGridEndeudamiento.MaxRows = vGridEndeudamiento.MaxRows - 1
   If vGridEndeudamiento.MaxRows = 0 Then vGridEndeudamiento.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbHistorial_Borrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGridHistorial.Row = vGridHistorial.ActiveRow
   vGridHistorial.Col = 1
   
   strSQL = "delete Crd_Clasificacion_Historial where cod_historial = '" & vGridHistorial.Text & "'"
   Call ConectionExecute(strSQL)
   
   vGridHistorial.Col = 1
   Call Bitacora("Elimina", "Clasificacion Historial : " & vGridHistorial.Text)
   
   vGridHistorial.DeleteRows vGridHistorial.ActiveRow, 1
   vGridHistorial.MaxRows = vGridHistorial.MaxRows - 1
   If vGridHistorial.MaxRows = 0 Then vGridHistorial.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswGarantia_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert Crd_clasificacion_Garantia_DT(cod_garantia,garantia) values('" & cboGarantia.ItemData(cboGarantia.ListIndex) _
          & "','" & Item.Tag & "')"
Else
   strSQL = "delete Crd_clasificacion_Garantia_DT where cod_Garantia = '" & cboGarantia.ItemData(cboGarantia.ListIndex) _
          & "' and garantia = '" & Item.Tag & "'"
End If
Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Me.MousePointer = vbHourglass


'Actualiza las Razones para los Grids
If Item.Index > 0 Then
   Call sbRazones_Cbos_Load
End If

 Select Case Item.Index
   Case 0 'Razones
      Call sbRazones_Load
   Case 1 'Garantia
      Call sbGarantia_Load
   Case 2 'Morosidad
      Call sbMora_Load
   Case 3 'Capacidad
      Call sbCapacidad_Load
   Case 4 'Endeudamiento
      Call sbEndeudamiento_Load
   Case 5 'Historial
      Call sbHistorial_Load
 End Select

Me.MousePointer = vbDefault

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxRazones_Guardar
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

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbRazones_Borrar
End If

End Sub


Private Sub vGridGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridGarantia.ActiveCol = vGridGarantia.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGarantia_Guardar
  If i = 0 Then Exit Sub
  vGridGarantia.Row = vGridGarantia.ActiveRow
  If vGridGarantia.MaxRows <= vGridGarantia.ActiveRow Then
    vGridGarantia.MaxRows = vGridGarantia.MaxRows + 1
    vGridGarantia.Row = vGridGarantia.MaxRows
    
    vGridGarantia.Col = vGridGarantia.MaxCols
    vGridGarantia.CellType = CellTypeComboBox
    vGridGarantia.TypeComboBoxList = mRazones
    vGridGarantia.TypeComboBoxEditable = False

  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridGarantia.MaxRows = vGridGarantia.MaxRows + 1
    vGridGarantia.InsertRows vGridGarantia.ActiveRow, 1
    vGridGarantia.Row = vGridGarantia.ActiveRow
    
    vGridGarantia.Col = vGridGarantia.MaxCols
    vGridGarantia.CellType = CellTypeComboBox
    vGridGarantia.TypeComboBoxList = mRazones
    vGridGarantia.TypeComboBoxEditable = False

End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbGarantia_Borrar
End If

End Sub


Private Sub vGridMora_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridMora.ActiveCol = vGridMora.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxMora_Guardar
  If i = 0 Then Exit Sub
  vGridMora.Row = vGridMora.ActiveRow
  If vGridMora.MaxRows <= vGridMora.ActiveRow Then
    vGridMora.MaxRows = vGridMora.MaxRows + 1
    vGridMora.Row = vGridMora.MaxRows
    
    vGridMora.Col = vGridMora.MaxCols
    vGridMora.CellType = CellTypeComboBox
    vGridMora.TypeComboBoxList = mRazones
    vGridMora.TypeComboBoxEditable = False

  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridMora.MaxRows = vGridMora.MaxRows + 1
    vGridMora.InsertRows vGridMora.ActiveRow, 1
    vGridMora.Row = vGridMora.ActiveRow
    
    vGridMora.Col = vGridMora.MaxCols
    vGridMora.CellType = CellTypeComboBox
    vGridMora.TypeComboBoxList = mRazones
    vGridMora.TypeComboBoxEditable = False

End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbMora_Borrar
End If

End Sub

Private Sub vGridCapacidad_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridCapacidad.ActiveCol = vGridCapacidad.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxCapacidad_Guardar
  If i = 0 Then Exit Sub
  vGridCapacidad.Row = vGridCapacidad.ActiveRow
  If vGridCapacidad.MaxRows <= vGridCapacidad.ActiveRow Then
    vGridCapacidad.MaxRows = vGridCapacidad.MaxRows + 1
    vGridCapacidad.Row = vGridCapacidad.MaxRows
    
    vGridCapacidad.Col = vGridCapacidad.MaxCols
    vGridCapacidad.CellType = CellTypeComboBox
    vGridCapacidad.TypeComboBoxList = mRazones
    vGridCapacidad.TypeComboBoxEditable = False

  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridCapacidad.MaxRows = vGridCapacidad.MaxRows + 1
    vGridCapacidad.InsertRows vGridCapacidad.ActiveRow, 1
    vGridCapacidad.Row = vGridCapacidad.ActiveRow
    
    vGridCapacidad.Col = vGridCapacidad.MaxCols
    vGridCapacidad.CellType = CellTypeComboBox
    vGridCapacidad.TypeComboBoxList = mRazones
    vGridCapacidad.TypeComboBoxEditable = False

End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbCapacidad_Borrar
End If

End Sub


Private Sub vGridEndeudamiento_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridEndeudamiento.ActiveCol = vGridEndeudamiento.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxEndeudamiento_Guardar
  If i = 0 Then Exit Sub
  vGridEndeudamiento.Row = vGridEndeudamiento.ActiveRow
  If vGridEndeudamiento.MaxRows <= vGridEndeudamiento.ActiveRow Then
    vGridEndeudamiento.MaxRows = vGridEndeudamiento.MaxRows + 1
    vGridEndeudamiento.Row = vGridEndeudamiento.MaxRows
    
    vGridEndeudamiento.Col = vGridEndeudamiento.MaxCols
    vGridEndeudamiento.CellType = CellTypeComboBox
    vGridEndeudamiento.TypeComboBoxList = mRazones
    vGridEndeudamiento.TypeComboBoxEditable = False

  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridEndeudamiento.MaxRows = vGridEndeudamiento.MaxRows + 1
    vGridEndeudamiento.InsertRows vGridEndeudamiento.ActiveRow, 1
    vGridEndeudamiento.Row = vGridEndeudamiento.ActiveRow
    
    vGridEndeudamiento.Col = vGridEndeudamiento.MaxCols
    vGridEndeudamiento.CellType = CellTypeComboBox
    vGridEndeudamiento.TypeComboBoxList = mRazones
    vGridEndeudamiento.TypeComboBoxEditable = False

End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbEndeudamiento_Borrar
End If

End Sub



Private Sub vGridHistorial_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer



If vGridHistorial.ActiveCol = vGridHistorial.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxHistorial_Guardar
  If i = 0 Then Exit Sub
  vGridHistorial.Row = vGridHistorial.ActiveRow
  If vGridHistorial.MaxRows <= vGridHistorial.ActiveRow Then
    vGridHistorial.MaxRows = vGridHistorial.MaxRows + 1
    vGridHistorial.Row = vGridHistorial.MaxRows
    
    vGridHistorial.Col = vGridHistorial.MaxCols
    vGridHistorial.CellType = CellTypeComboBox
    vGridHistorial.TypeComboBoxList = mRazones
    vGridHistorial.TypeComboBoxEditable = False

  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridHistorial.MaxRows = vGridHistorial.MaxRows + 1
    vGridHistorial.InsertRows vGridHistorial.ActiveRow, 1
    vGridHistorial.Row = vGridHistorial.ActiveRow
    
    vGridHistorial.Col = vGridHistorial.MaxCols
    vGridHistorial.CellType = CellTypeComboBox
    vGridHistorial.TypeComboBoxList = mRazones
    vGridHistorial.TypeComboBoxEditable = False

End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbHistorial_Borrar
End If

End Sub

