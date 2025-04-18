VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmIVR_Cat_Recursos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "SCGI Fuente de Recursos"
   ClientHeight    =   6960
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9564
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9564
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5532
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   9132
      _Version        =   524288
      _ExtentX        =   16108
      _ExtentY        =   9758
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
      MaxCols         =   483
      ScrollBars      =   2
      SpreadDesigner  =   "frmIVR_Cat_Recursos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   7212
      _Version        =   1310720
      _ExtentX        =   12721
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Fuentes de Recursos"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmIVR_Cat_Recursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub sbConsulta()
Dim strSQL As String


vPaso = True

strSQL = "select COD_RECURSO,descripcion,Activa,0 from IVR_FUENTE_RECURSOS" _
      & " order by COD_RECURSO"
Call sbCargaGrid(vGrid, 4, strSQL)

vPaso = False

End Sub

Private Sub Form_Load()

vModulo = 22

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbConsulta

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

strSQL = "select isnull(count(*),0) as Existe from IVR_FUENTE_RECURSOS " _
       & " where COD_RECURSO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into IVR_FUENTE_RECURSOS(COD_RECURSO,DESCRIPCION,ACTIVA, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Fuente de Recursos:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update IVR_FUENTE_RECURSOS set descripcion = '" & vGrid.Text & "',Activa = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where COD_RECURSO = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Fuente de Recursos:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If Col = 4 Then
   vGrid.Row = Row
   vGrid.Col = 1
   
   gIVR_Cuentas.Tipo = "R"
   gIVR_Cuentas.Codigo_1 = vGrid.Text
   gIVR_Cuentas.Codigo_2 = ""
   
   vGrid.Col = 2
   gIVR_Cuentas.Descripcion = vGrid.Text
   
   frmIVR_Cat_Cuentas_Contables.Show vbModal
   

End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If vGrid.ActiveCol = vGrid.MaxCols - 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete IVR_FUENTE_RECURSOS where COD_RECURSO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Fuente de Recursos:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub









