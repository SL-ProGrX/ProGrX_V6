VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_TablaDevoluciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Devoluciones (Prioridad del proceso de aplicación)"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   9015
      _Version        =   524288
      _ExtentX        =   15901
      _ExtentY        =   5530
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   486
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_TablaDevoluciones.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFSL_TablaDevoluciones.frx":06D6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   7695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Devoluciones"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmFSL_TablaDevoluciones.frx":0792
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmFSL_TablaDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUltimaSeleccion As String, mLista As String

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 22

vGrid.AppearanceStyle = fxGridStyle

strUltimaSeleccion = ""
mLista = ""
Call sbListaGarantias

strSQL = "select Fsl.COD_DEVOLUCION,Fsl.Fecha_Inicio,Fsl.Fecha_Corte" _
       & ", rtrim(Fsl.Garantia) + ' - ' + rtrim(Gar.descripcion) as 'GARANTIA'" _
       & ",case when BASE_APLICACION = 'S' then 'Saldo' else 'Formalizado' end as 'Base'" _
       & ",Fsl.Porcentaje, Fsl.Registro_Fecha,Fsl.Registro_Usuario" _
       & " from FSL_TABLA_DEVOLUCIONES Fsl inner join CRD_Garantia_Tipos Gar on Fsl.Garantia = Gar.Garantia" _
       & " order by Fsl.Fecha_Inicio"
Call sbCargaGridLocal(vGrid, 6, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbListaGarantias()
Dim rs As New ADODB.Recordset, strSQL As String


strSQL = "select rtrim(Garantia) + ' - ' + rtrim(descripcion) as 'GARANTIA'" _
       & " from CRD_Garantia_Tipos" _
       & " order by Garantia"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF And strUltimaSeleccion = "" Then
 strUltimaSeleccion = rs!Garantia
End If

mLista = ""

Do While Not rs.EOF
  If Len(mLista) = 0 Then
    mLista = Chr$(9) & rs!Garantia
  Else
    mLista = mLista & Chr$(9) & rs!Garantia
  End If
  rs.MoveNext
Loop
rs.Close

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenForwardOnly


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!COD_DEVOLUCION)
     Case 2
        vGrid.Text = CStr(rs!Fecha_Inicio)
     Case 3
        vGrid.Text = CStr(rs!Fecha_Corte)
     Case 4
        vGrid.CellType = CellTypeComboBox
        vGrid.TypeComboBoxList = mLista
        vGrid.TypeComboBoxEditable = False
        vGrid.Text = rs!Garantia
        strUltimaSeleccion = rs!Garantia
     Case 5
        vGrid.Text = rs!Base
     Case 6
        vGrid.Text = CStr(rs!Porcentaje)
        
    End Select
  
  Next i
  
    vGrid.Col = 6
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario " & IIf(IsNull(rs!Registro_Usuario), "...!", rs!Registro_Usuario) _
                     & vbCrLf & "Fecha " & IIf(IsNull(rs!registro_Fecha), "...!", rs!registro_Fecha)
                     
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 4
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mLista
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = strUltimaSeleccion
  
Me.MousePointer = vbDefault

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCodigo As Long

On Error GoTo vError

vGrid.Row = vGrid.ActiveRow

vGrid.Col = 1
fxGuardar = 0

If vGrid.Text = "" Then 'Insertar
  
  strSQL = "select coalesce(max(COD_DEVOLUCION),0) + 1 as Ultimo from FSL_TABLA_DEVOLUCIONES"
  rs.Open strSQL, glogon.Conection, adOpenStatic
      pCodigo = rs!Ultimo
  rs.Close
  
  
  strSQL = "insert into FSL_TABLA_DEVOLUCIONES(COD_DEVOLUCION,Fecha_Inicio,Fecha_Corte,Garantia,Base_Aplicacion,Porcentaje" _
         & ",registro_fecha,registro_usuario) values(" & pCodigo & ",'"
  vGrid.Col = 2
  strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & "','"
  vGrid.Col = 3
  strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & "','"
  vGrid.Col = 4
  strSQL = strSQL & SIFGlobal.fxSIFCodText(vGrid.Text) & "','"
  vGrid.Col = 5
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  vGrid.Col = 6
  strSQL = strSQL & CCur(vGrid.Text) & ", getdate(),'" & glogon.Usuario & "')"
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  vGrid.Text = CStr(pCodigo)
  
  Call Bitacora("Registra", "Tabla Devolución Id: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update FSL_TABLA_DEVOLUCIONES set Fecha_Inicio = '" & Format(vGrid.Text, "yyyy/mm/dd") & "', Fecha_Corte = '"
    vGrid.Col = 3
    strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & "', Garantia = '"
    vGrid.Col = 4
    strSQL = strSQL & SIFGlobal.fxSIFCodText(vGrid.Text) & "', Base_Aplicacion = '"
    vGrid.Col = 5
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', Porcentaje = "
    vGrid.Col = 6
    strSQL = strSQL & CCur(vGrid.Text)
    vGrid.Col = 1
    strSQL = strSQL & " where COD_DEVOLUCION = " & vGrid.Text
    glogon.Conection.Execute strSQL
 
    vGrid.Col = 1
   
    Call Bitacora("Modifica", "Tabla Devolución Id: " & vGrid.Text)
 
End If

fxGuardar = 1

vGrid.Col = 4
strUltimaSeleccion = vGrid.Text

vGrid.Col = 6
vGrid.TextTip = TextTipFixed
vGrid.TextTipDelay = 1000
                
vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
vGrid.CellNote = "Usuario " & glogon.Usuario _
                 & vbCrLf & " Fecha " & Format(Date, "dd/mm/yyyy")
                 
Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Long


On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    vGrid.Col = 4
    vGrid.CellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mLista
    vGrid.TypeComboBoxEditable = False
    vGrid.Text = strUltimaSeleccion
         
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 4
    vGrid.CellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mLista
    vGrid.TypeComboBoxEditable = False
    vGrid.Text = strUltimaSeleccion
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        strSQL = "delete FSL_TABLA_DEVOLUCIONES where COD_DEVOLUCION = " & vGrid.Text
        glogon.Conection.Execute strSQL
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tabla Devolución Id: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


