VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_Sectores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sectores de Mercado"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5292
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6612
      _Version        =   524288
      _ExtentX        =   11663
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
      SpreadDesigner  =   "frmAF_Sectores.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mercado: Sectores"
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
      Height          =   495
      Index           =   2
      Left            =   1880
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_Sectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verificar y establecer permisos sobre el formulario.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

vModulo = 1
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
If vGrid.Text = "" Then
    vGrid.Col = 2
    strSQL = "insert afi_sectores(descripcion) values('" & vGrid.Text & "')"
    Call ConectionExecute(strSQL)
  
    Call Bitacora("Registra", "Sector : " & vGrid.Text)
  
  
    vGrid.Col = 2
    strSQL = "select max(cod_sector) as ultimo from afi_sectores where descripcion = '" & vGrid.Text & "'"
    Call OpenRecordSet(rs, strSQL)
      vGrid.Col = 1
      vGrid.Text = CStr(rs!ultimo)
    rs.Close
   
   Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update afi_sectores set descripcion = '" & vGrid.Text _
           & "' where cod_sector = "
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text
    Call ConectionExecute(strSQL)
    
   End If

   vGrid.Col = 1
   fxGuardar = vGrid.Text
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

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


'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1

       If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
       
               vGrid.Row = vGrid.ActiveRow
               vGrid.Col = 1
               strSQL = "delete afi_sectores where cod_sector= " & vGrid.Text
               Call ConectionExecute(strSQL)
               
               strSQL = vGrid.Text
               vGrid.Col = 2
               Call Bitacora("Elimina", "Sector : " & strSQL & " - " & vGrid.Text)
               vGrid.Col = 1
               
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If
End If

Exit Sub

vError:

  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Cargar en el Objeto adoData Todas las causas de renuncia que existen.
'REFERENCIAS:   sbToolBarIconos - (Carga los iconos para la barra de herramientas)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String

On Error GoTo error
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


strSQL = "select cod_sector,descripcion from afi_sectores order by cod_sector"
Call sbCargaGrid(vGrid, 2, strSQL)

Exit Sub
error:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





