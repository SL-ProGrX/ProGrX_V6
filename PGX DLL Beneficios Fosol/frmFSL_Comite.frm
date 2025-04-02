VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Comite 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comités de FOSOL"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Comités"
      TabPicture(0)   =   "frmFSL_Comite.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Miembros"
      TabPicture(1)   =   "frmFSL_Comite.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboComite"
      Tab(1).Control(1)=   "vGridMiembros"
      Tab(1).Control(2)=   "Label6"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox cboComite 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   6135
      End
      Begin FPSpreadADO.fpSpread vGridMiembros 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   3
         Top             =   960
         Width           =   12015
         _Version        =   524288
         _ExtentX        =   21193
         _ExtentY        =   6376
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   6
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Comite.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4095
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   9615
         _Version        =   524288
         _ExtentX        =   16960
         _ExtentY        =   7223
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Comite.frx":0791
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Comité"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   11160
      X2              =   0
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmFSL_Comite.frx":0E08
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cómite de Evaluacion.: Casos FOSOL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   6975
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
vModulo = 22
End Sub



Private Sub cboComite_Click()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select CEDULA,Nombre,USUARIO_VINCULADO,Registro_Fecha,Salida_Fecha,Activo" _
       & " from FSL_COMITES_MIEMBROS" _
       & " where COD_COMITE = '" & SIFGlobal.fxSIFCodText(cboComite.Text) _
       & "' order by Activo desc, CEDULA"
Call sbCargaGrid(vGridMiembros, 6, strSQL)
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 22
vGrid.AppearanceStyle = fxGridStyle

ssTab.Tab = 0
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
rs.Open strSQL, glogon.Conection, adOpenStatic


If rs!Existe = 0 Then

   
  strSQL = "insert " & pTabla & "(" & pCodigo & " ,Descripcion, Numero_Resolutores, Activo,registro_fecha,registro_usuario) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",getdate(),'" & glogon.Usuario & "')"
  
  glogon.Conection.Execute strSQL

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
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Modifica", "Comité de FOSOL Id.:" & vGrid.Text)

End If

rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

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
           & " where " & pCodigo & " = '" & .Text & "' AND COD_COMITE = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    
    If rs!Existe = 0 Then
    
       
      strSQL = "insert " & pTabla & "(" & pCodigo & ",COD_COMITE, Nombre, USUARIO_VINCULADO,Activo,registro_fecha,registro_usuario) values('" _
             & .Text & "','" & SIFGlobal.fxSIFCodText(cboComite.Text) & "','"
      .Col = 2
      strSQL = strSQL & .Text & "','"
      .Col = 3
      strSQL = strSQL & .Text & "',"
      .Col = 6
      strSQL = strSQL & .Value & ",getdate(),'" & glogon.Usuario & "')"
      
      glogon.Conection.Execute strSQL
    
      .Col = 1
      
      Call Bitacora("Registra", "Comité Miembro: " & SIFGlobal.fxSIFCodText(cboComite.Text) & ".. Id.:" & .Text)
    
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
            strSQL = strSQL & .Text & "' and COD_COMITE = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
      Else
            .Col = 2
            strSQL = "update " & pTabla & " set Nombre = '" & .Text & "', Salida_Fecha = getdate(), Salida_Usuario = '" _
                   & glogon.Usuario & "', Usuario_Vinculado = '"
            .Col = 3
            strSQL = strSQL & .Text & "',  Activo = "
            .Col = 6
            strSQL = strSQL & .Value & " where " & pCodigo & " = '"
            .Col = 1
            strSQL = strSQL & .Text & "' and COD_COMITE = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
      End If
      glogon.Conection.Execute strSQL
    
    
      .Col = 1
      Call Bitacora("Modifica", "Comité Miembro: " & SIFGlobal.fxSIFCodText(cboComite.Text) & ".. Id.:" & .Text)
    
    End If

End With

rs.Close

fxGuardarMiembro = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function






Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

If ssTab.Tab = 1 Then

    vPaso = True
        strSQL = "select RTRIM(COD_COMITE) + ' - ' + DESCRIPCION as ItmX FROM FSL_COMITES WHERE ACTIVO = 1"
        Call sbLlenaCbo(cboComite, strSQL, False, False)
    vPaso = False
    Call cboComite_Click
End If

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
        glogon.Conection.Execute strSQL

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
  MsgBox Err.Description, vbCritical

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
        strSQL = "delete FSL_COMITES_MIEMBROS where COD_COMITE = '" & SIFGlobal.fxSIFCodText(cboComite.Text) _
                & "' AND CEDULA = '" & vGridMiembros.Text & "'"
        glogon.Conection.Execute strSQL

        strSQL = vGridMiembros.Text
        vGridMiembros.Col = 1
        Call Bitacora("Elimina", "Comité Miembro: " & SIFGlobal.fxSIFCodText(cboComite.Text) & " .. Id.:" & vGridMiembros.Text)

        vGridMiembros.DeleteRows vGridMiembros.ActiveRow, 1
        vGridMiembros.MaxRows = vGridMiembros.MaxRows - 1
        vGridMiembros.Row = vGridMiembros.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub




