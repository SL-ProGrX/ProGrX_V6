VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPreaTablaImpRenta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla : Cálculo Impuesto de Renta"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2652
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   5532
      _Version        =   524288
      _ExtentX        =   9758
      _ExtentY        =   4678
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
      SpreadDesigner  =   "frmPreaTablaImpRenta.frx":0000
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla para Renta sobre Salario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaTablaImpRenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub sbGrid_Load()

On Error GoTo vError

strSQL = "select idx,desde,hasta,porcentaje from crd_prea_tabla_impuesto order by desde"
Call sbCargaGrid(vGrid, 4, strSQL)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 3 'Modulo de Credito

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)
 
Call sbGrid_Load

End Sub

Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If vGrid.Text = "" Or vGrid.Text = "0" Then
   vGrid.Col = 2
   strSQL = "insert into Crd_Prea_Tabla_Impuesto(desde,hasta,porcentaje, REGISTRO_USUARIO, REGISTRO_FECHA)" _
          & " values(" & CCur(vGrid.Text) & ","
   vGrid.Col = 3
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.Col = 4
   strSQL = strSQL & CCur(vGrid.Text) & ",'" & glogon.Usuario & "', getdate() )"
   
   Call ConectionExecute(strSQL)
    
    strSQL = "select isnull(max(IDx),0) as ultimo from Crd_Prea_Tabla_Impuesto"
    Call OpenRecordSet(rs, strSQL)
      vGrid.Col = 1
      vGrid.Text = CStr(rs!ultimo)
    rs.Close
   
    Call Bitacora("Registra", "Rango Estudio de Credito - Impuesto Renta [ID]: " & vGrid.Text)
    
    MsgBox "Impuesto de Renta Id: " & vGrid.Text & ", Registrado satisfactoriamente!", vbInformation
   
   Else 'Actualizar
    vGrid.Col = 2
    strSQL = "update Crd_Prea_Tabla_Impuesto set desde = " & CCur(vGrid.Text)
    vGrid.Col = 3
    strSQL = strSQL & ", hasta = " & CCur(vGrid.Text)
    vGrid.Col = 4
    strSQL = strSQL & ", porcentaje = " & CCur(vGrid.Text) & ", modifica_Usuario = '" & glogon.Usuario & "', modifica_fecha = getdate()"
    vGrid.Col = 1
    strSQL = strSQL & " where Idx = " & vGrid.Text
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Rango Estudio de Credito - Impuesto Renta [ID]: " & vGrid.Text)
    
    MsgBox "Impuesto de Renta Id: " & vGrid.Text & ", Modificado satisfactoriamente!", vbInformation
    
   End If

   vGrid.Col = 1
   fxGuardar = vGrid.Text
   
   Exit Function
   
vError:
 fxGuardar = 0
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
    Call sbBorrar
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If vGrid.Text = "" Then
    Exit Sub
End If

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   strSQL = "delete Crd_Prea_Tabla_Impuesto where IDx = " & vGrid.Text
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Elimina", "Rango Estudio de Credito - Impuesto Renta [ID]: " & vGrid.Text)
    
   MsgBox "Impuesto de Renta Id: " & vGrid.Text & ", Eliminado!", vbInformation
   
   Call sbGrid_Load
End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

