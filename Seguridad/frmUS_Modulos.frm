VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmUS_Modulos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulos del Sistema"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   HelpContextID   =   24
   Icon            =   "frmUS_Modulos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9375
      _Version        =   524288
      _ExtentX        =   16536
      _ExtentY        =   11668
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
      MaxCols         =   486
      ScrollBars      =   2
      SpreadDesigner  =   "frmUS_Modulos.frx":000C
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Módulos del Sistema"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmUS_Modulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String


vModulo = 13
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


strSQL = "select * from us_Modulos order by modulo"
Call sbCargaGrid(vGrid, 4, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Boolean
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then Exit Function
  
strSQL = "select count(*) as 'Existe' from us_Modulos where modulo = " & vGrid.Text
Call OpenRecordSet(rs, strSQL, 1)
   
If rs!Existe = 0 Then
  
  
    strSQL = "insert into us_Modulos(modulo,nombre,descripcion,activo) values(" & vGrid.Text & ",'"
    vGrid.Col = 2
    strSQL = strSQL & Trim(UCase(vGrid.Text)) & "','"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Value & ")"
    
      
    Call ConectionExecute(strSQL, 1)
  
    vGrid.Col = 1
 
    Call Bitacora("Registra", "Modulo del Sistema: " & vGrid.Text)
  
 Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update us_Modulos set nombre = '" & Trim(UCase(vGrid.Text)) & "',descripcion = '"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & "',Activo = "
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Value & " Where modulo = "
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text
    Call ConectionExecute(strSQL, 1)
    
    Call Bitacora("Modifica", "Modulo del Sistema: " & vGrid.Text)
    
 End If

rs.Close

vGrid.Col = 1
fxGuardar = vGrid.Text
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  vGrid.Text = i
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
        
        strSQL = "delete US_Opciones where Modulo = " & vGrid.Text
        Call ConectionExecute(strSQL, 1)
    
        strSQL = "delete US_Formularios where Modulo = " & vGrid.Text
        Call ConectionExecute(strSQL, 1)
        
        strSQL = "delete US_Modulos where Modulo = " & vGrid.Text
        Call ConectionExecute(strSQL, 1)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Módulo del Sistema: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If


End Sub



