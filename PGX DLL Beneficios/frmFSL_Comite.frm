VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmFSL_Comite 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comités de FOSOL"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   12135
      _Version        =   1441793
      _ExtentX        =   21405
      _ExtentY        =   11033
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
      Item(0).Caption =   "Comités"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Miembros"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "cboComite"
      Item(1).Control(1)=   "vGridMiembros"
      Item(1).Control(2)=   "Label6"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5775
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   9615
         _Version        =   524288
         _ExtentX        =   16960
         _ExtentY        =   10186
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Comite.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridMiembros 
         Height          =   5055
         Left            =   -70000
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   12015
         _Version        =   524288
         _ExtentX        =   21193
         _ExtentY        =   8916
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   6
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Comite.frx":069F
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboComite 
         Height          =   330
         Left            =   -64360
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   6135
         _Version        =   1441793
         _ExtentX        =   10821
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Comité"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66400
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cómite de Evaluacion.: Casos FOSOL"
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
      Height          =   492
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmFSL_Comite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub Form_Activate()
vModulo = 7
End Sub



Private Sub cboComite_Click()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select CEDULA,Nombre,USUARIO_VINCULADO,Registro_Fecha,Salida_Fecha,Activo" _
       & " from FSL_COMITES_MIEMBROS" _
       & " where COD_COMITE = '" & cboComite.ItemData(cboComite.ListIndex) _
       & "' order by Activo desc, CEDULA"
Call sbCargaGrid(vGridMiembros, 6, strSQL)

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 7
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


tcMain.Item(0).Selected = True


strSQL = "select COD_COMITE,descripcion,Numero_Resolutores,Activo from FSL_COMITES order by COD_COMITE"
Call sbCargaGrid(vGrid, 4, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pExiste As Long, pCodigo As String, pTabla As String

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

pCodigo = "COD_COMITE"
pTabla = "FSL_COMITES"

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then Exit Function

strSQL = "select isnull(count(*),0) as Existe from " & pTabla _
       & " where " & pCodigo & " = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)




If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & " ,Descripcion, Numero_Resolutores, Activo,registro_fecha,registro_usuario) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Comité de FOSOL Id.:" & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update " & pTabla & " set Descripcion = '" & vGrid.Text & "', Numero_Resolutores = "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & ", Activo = "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & " where " & pCodigo & " = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Comité de FOSOL Id.:" & vGrid.Text)

End If

rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Function


Private Function fxGuardarMiembro() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pExiste As Long, pCodigo As String, pTabla As String

On Error GoTo vError

pCodigo = "CEDULA"
pTabla = "FSL_COMITES_MIEMBROS"

fxGuardarMiembro = 0


With vGridMiembros

    .Row = .ActiveRow
    .Col = 1
     
    If Trim(.Text) = "" Then Exit Function
    
    strSQL = "select isnull(count(*),0) as Existe from " & pTabla _
           & " where " & pCodigo & " = '" & .Text & "' AND COD_COMITE = '" & cboComite.ItemData(cboComite.ListIndex) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    
    If rs!Existe = 0 Then
    
       
      strSQL = "insert " & pTabla & "(" & pCodigo & ",COD_COMITE, Nombre, USUARIO_VINCULADO,Activo,registro_fecha,registro_usuario) values('" _
             & .Text & "','" & cboComite.ItemData(cboComite.ListIndex) & "','"
      .Col = 2
      strSQL = strSQL & .Text & "','"
      .Col = 3
      strSQL = strSQL & .Text & "',"
      .Col = 6
      strSQL = strSQL & .Value & ",getdate(),'" & glogon.Usuario & "')"
      
      Call ConectionExecute(strSQL)
    
      .Col = 1
      
      Call Bitacora("Registra", "Comité Miembro: " & cboComite.ItemData(cboComite.ListIndex) & ".. Id.:" & .Text)
    
    Else 'Actualizar
      
      .Col = 6
      If .Value = vbChecked Then
            .Col = 2
            strSQL = "update " & pTabla & " set Nombre = '" & .Text & "', Usuario_Vinculado = '"
            .Col = 3
            strSQL = strSQL & .Text & "',  Activo = "
            .Col = 6
            strSQL = strSQL & .Value & " where " & pCodigo & " = '"
            .Col = 1
            strSQL = strSQL & .Text & "' and COD_COMITE = '" & cboComite.ItemData(cboComite.ListIndex) & "'"
      Else
            .Col = 2
            strSQL = "update " & pTabla & " set Nombre = '" & .Text & "', Salida_Fecha = getdate(), Salida_Usuario = '" _
                   & glogon.Usuario & "', Usuario_Vinculado = '"
            .Col = 3
            strSQL = strSQL & .Text & "',  Activo = "
            .Col = 6
            strSQL = strSQL & .Value & " where " & pCodigo & " = '"
            .Col = 1
            strSQL = strSQL & .Text & "' and COD_COMITE = '" & cboComite.ItemData(cboComite.ListIndex) & "'"
      End If
      Call ConectionExecute(strSQL)
    
    
      .Col = 1
      Call Bitacora("Modifica", "Comité Miembro: " & cboComite.ItemData(cboComite.ListIndex) & ".. Id.:" & .Text)
    
    End If

End With

rs.Close

fxGuardarMiembro = 1

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
        strSQL = "delete FSL_COMITES where COD_COMITE = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Comité de FOSOL Id.:" & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 1 Then

    vPaso = True
        strSQL = "select RTRIM(COD_COMITE) as 'IdX', rtrim(DESCRIPCION) as 'ItmX' FROM FSL_COMITES WHERE ACTIVO = 1"
        Call sbCbo_Llena_New(cboComite, strSQL, False, True)
    vPaso = False
    Call cboComite_Click
End If

End Sub

Private Sub vGridMiembros_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGridMiembros.ActiveCol = vGridMiembros.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarMiembro
  If i = 0 Then Exit Sub
  vGridMiembros.Row = vGridMiembros.ActiveRow
  If vGridMiembros.MaxRows <= vGridMiembros.ActiveRow Then
    vGridMiembros.MaxRows = vGridMiembros.MaxRows + 1
    vGridMiembros.Row = vGridMiembros.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridMiembros.MaxRows = vGridMiembros.MaxRows + 1
    vGridMiembros.InsertRows vGridMiembros.ActiveRow, 1
    vGridMiembros.Row = vGridMiembros.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGridMiembros.Row = vGridMiembros.ActiveRow
        vGridMiembros.Col = 1
        strSQL = "delete FSL_COMITES_MIEMBROS where COD_COMITE = '" & cboComite.ItemData(cboComite.ListIndex) _
                & "' AND CEDULA = '" & vGridMiembros.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGridMiembros.Text
        vGridMiembros.Col = 1
        Call Bitacora("Elimina", "Comité Miembro: " & cboComite.ItemData(cboComite.ListIndex) & " .. Id.:" & vGridMiembros.Text)

        vGridMiembros.DeleteRows vGridMiembros.ActiveRow, 1
        vGridMiembros.MaxRows = vGridMiembros.MaxRows - 1
        vGridMiembros.Row = vGridMiembros.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
