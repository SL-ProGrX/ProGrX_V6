VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_BeneficioProd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beneficios / Productos (Beneficios en Especies)"
   ClientHeight    =   6984
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8556
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6984
   ScaleWidth      =   8556
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8172
      _Version        =   524288
      _ExtentX        =   14414
      _ExtentY        =   9546
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
      MaxCols         =   498
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_BeneficioProd.frx":0000
      VScrollSpecialType=   2
      Appearance      =   1
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Catálogo de Productos y Servicios"
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
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmAF_BeneficioProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select cod_producto,descripcion,costo_unidad from afi_bene_productos" _
      & " order by cod_producto"
Call sbCargaGrid(vGrid, 3, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from afi_bene_productos " _
       & " where cod_producto = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into afi_bene_productos(cod_producto,descripcion,costo_unidad) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.Col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ")"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Producto Beneficio : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update afi_bene_productos set descripcion = '" & vGrid.Text & "',costo_unidad = "
 vGrid.Col = 3
 strSQL = strSQL & CCur(vGrid.Text) & " where cod_producto = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Producto Beneficio : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete afi_bene_productos where cod_producto = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Producto de Beneficio : " & vGrid.Text)
        
        vGrid.Col = 1
        strSQL = "select cod_producto,descripcion,costo_unidad from afi_bene_productos" _
              & " order by cod_producto"
        Call sbCargaGrid(vGrid, 3, strSQL)
     End If
  
  Case "REPORTES"
'     Call sbReportes("Caracteristicas", Me)

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Elimina
If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        strSQL = "delete afi_bene_productos where cod_producto = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Producto de Beneficio : " & vGrid.Text)
        
        vGrid.Col = 1
        strSQL = "select cod_producto,descripcion,costo_unidad from afi_bene_productos" _
              & " order by cod_producto"
        Call sbCargaGrid(vGrid, 3, strSQL)
     End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub






