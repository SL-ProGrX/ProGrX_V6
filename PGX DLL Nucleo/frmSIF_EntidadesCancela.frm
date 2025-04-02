VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSIF_EntidadesCancela 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de Entidades que Cancelan Operaciones Internas"
   ClientHeight    =   6408
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   8388
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6408
   ScaleWidth      =   8388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5292
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   7932
      _Version        =   524288
      _ExtentX        =   13991
      _ExtentY        =   9335
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
      SpreadDesigner  =   "frmSIF_EntidadesCancela.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entidades que Cancelan (Op. Crd. Int.)"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   6492
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSIF_EntidadesCancela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 10
 
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select COD_ENTIDAD_PAGO,descripcion,activa from SIF_ENTIDADES_PAGO"
Call sbCargaGrid(vGrid, 3, strSQL)
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
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


If KeyCode = vbKeyDelete Then
   'Aqui codigo de Borrado
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   If Trim(vGrid.Text) <> "" Then
    strSQL = "Delete SIF_ENTIDADES_PAGO where COD_ENTIDAD_PAGO =  '" & UCase(vGrid.Text) & "'"
    Call ConectionExecute(strSQL)
        
    Call Bitacora("Elimina", "Entidades Pagadoras: " & vGrid.Text)
   End If
   
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
End If


If KeyCode = vbKeyInsert Then
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.InsertRows vGrid.ActiveRow, 1
  vGrid.Row = vGrid.ActiveRow
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If vGrid.Text = "" Then vGrid.Text = 0
strSQL = "select isnull(count(*),0) as Existe from SIF_ENTIDADES_PAGO  " _
       & " where COD_ENTIDAD_PAGO ='" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
    If Trim(vGrid.Text) = "" Then Exit Function
    strSQL = "insert into SIF_ENTIDADES_PAGO(COD_ENTIDAD_PAGO,descripcion,activa,registro_usuario,registro_fecha)" _
           & " values('" & UCase(vGrid.Text) & "',"
    vGrid.Col = 2
    strSQL = strSQL & "'" & UCase(vGrid.Text) & "',"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
    
    Call ConectionExecute(strSQL)
    
    vGrid.Col = 1
    Call Bitacora("Registra", "Entidades Pagadoras: " & vGrid.Text)

Else 'Actualizar
    
    vGrid.Col = 2
    strSQL = "update SIF_ENTIDADES_PAGO set descripcion= '" & UCase(vGrid.Text) & "',activa = "
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & " where COD_ENTIDAD_PAGO =  '"
    vGrid.Col = 1
    strSQL = strSQL & UCase(vGrid.Text) & "'"
    
     
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Entidades Pagadoras: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



