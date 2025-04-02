VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_CD_MiembrosJunta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Directores de Zonas"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510
   Icon            =   "frmAF_CD_MiembrosJunta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAF_CD_MiembrosJunta.frx":3482
   ScaleHeight     =   5250
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4125
      Left            =   60
      TabIndex        =   0
      Top             =   990
      Width           =   9315
      _Version        =   524288
      _ExtentX        =   16431
      _ExtentY        =   7276
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
      MaxCols         =   4
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_CD_MiembrosJunta.frx":9CD4
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Miembros de Junta Directiva (Directores de Zona)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   9360
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmAF_CD_MiembrosJunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
 vModulo = 23
End Sub

Private Sub Form_Load()
Dim strSQL As String

 vModulo = 23
 
strSQL = "select coalesce(cod_director,1),Nombre,puesto,Activo from afi_cd_directores "
Call sbCargaGrid(vGrid, 4, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Long, strSQL As String
Dim rs As New ADODB.Recordset
On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        If vGrid.MaxRows <= vGrid.ActiveRow Then
           vGrid.MaxRows = vGrid.MaxRows + 1
           vGrid.Row = vGrid.MaxRows
        End If
  End If 'Actualiza o Inserta
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 1

  If vGrid.Text = "" Then Exit Sub
     
     strSQL = "select cod_comite,descripcion,cod_director from afi_cd_comites " _
              & "where cod_director = " & vGrid.Text & ""
              rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     
          If Not rs.EOF Then
                   MsgBox "Actualmente este director pertenece al comité " & rs!cod_comite & " " & rs!Descripcion & " no podra eliminarlo", vbInformation, "Información"
                   rs.Close
                   Exit Sub
          Else
                  
                  i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
                  If i = vbYes Then
                    
                      strSQL = "delete afi_cd_directores where cod_director = " & vGrid.Text
                      glogon.Conection.Execute strSQL
                     
                      strSQL = vGrid.Text
                      vGrid.Col = 2
                      'Call Bitacora("Elimina", "Director: " & vGrid.Text & ")
                     
                      vGrid.DeleteRows vGrid.ActiveRow, 1
                      vGrid.MaxRows = vGrid.MaxRows - 1
                      If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
                     
          End If
          rs.Close
   End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub
Private Function fxConsec()
Dim strSQL As String, rs As New ADODB.Recordset

     
    strSQL = "select coalesce(max(cod_director),0) + 1 as Ultimo from afi_cd_directores"
    rs.Open strSQL, glogon.Conection, adOpenStatic
        fxConsec = rs!ultimo
    rs.Close

End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then
    
    vGrid.Col = 1
    strSQL = "insert afi_cd_directores(cod_director,nombre,puesto,activo) values(" & fxConsec & ",'"
    vGrid.Col = 2
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & "','"
    vGrid.Col = 3
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & "',"
    vGrid.Col = 4
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Value) & ")"
    
    glogon.Conection.Execute strSQL
    
    strSQL = "select coalesce(cod_director,1),Nombre,puesto,Activo from afi_cd_directores "
    Call sbCargaGrid(vGrid, 4, strSQL)
  
    vGrid.Col = 2
    'Call Bitacora("Registra", "Directores: " & vGrid.Text & " Ced: " & GLOBALES.gCedulaActual & " ID." & strSQL)
    fxGuardar = 1
   
   Else 'Actualizar
  
    vGrid.Col = 2
    strSQL = "update afi_cd_directores set nombre = '" & IIf((vGrid.Text = ""), 0, vGrid.Text) & "',puesto ='"
    vGrid.Col = 3
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & "',activo = "
    vGrid.Col = 4
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Value) & " "
    vGrid.Col = 1
    strSQL = strSQL & "where cod_director = " & vGrid.Text
    
    glogon.Conection.Execute strSQL
    
    strSQL = vGrid.Text
    
    vGrid.Col = 2
    'Call Bitacora("Modifica", "Directores: " & vGrid.Text & " ID: " & GLOBALES.gCedulaActual & " ID." & strSQL)
    
   End If
   


Exit Function
vError:
MsgBox Err.Description, vbCritical
fxGuardar = 0
End Function

