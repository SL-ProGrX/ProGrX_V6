VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmAF_DistribucionPolitica 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribución Política"
   ClientHeight    =   6960
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   12192
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   12192
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   5880
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5652
      _ExtentX        =   9970
      _ExtentY        =   10372
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgExplorer"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   4560
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":11B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":14D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":17F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":1B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":1E28
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":2144
            Key             =   "imgUsuario"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":2A20
            Key             =   "imgGrupo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":32FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":3618
            Key             =   "imgCuentas"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":3934
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":3A41
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":3B61
            Key             =   "imgFolder"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_DistribucionPolitica.frx":3C7D
            Key             =   "imgAsientos"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5892
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   6132
      _Version        =   524288
      _ExtentX        =   10816
      _ExtentY        =   10393
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
      SpreadDesigner  =   "frmAF_DistribucionPolitica.frx":3D8B
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Image imgRefrescar 
      Height          =   192
      Left            =   11760
      Picture         =   "frmAF_DistribucionPolitica.frx":42AC
      Top             =   240
      Width           =   192
   End
   Begin XtremeShortcutBar.ShortcutCaption lblY 
      Height          =   372
      Left            =   10680
      TabIndex        =   6
      Top             =   120
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblX 
      Height          =   372
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   4812
      _Version        =   1245187
      _ExtentX        =   8488
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitule 
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5772
      _Version        =   1245187
      _ExtentX        =   10181
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Distribución Política"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblNodeLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Canton"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Linea"
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label lblNodeLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Linea"
      Top             =   6600
      Width           =   2295
   End
End
Attribute VB_Name = "frmAF_DistribucionPolitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vModifica As String
Dim vCantonMascara As String, vDistritoMascara As String

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle


'Mascara del canton
vCantonMascara = "0"
strSQL = "select MAX(LEN(canton)) as Caracteres from CANTONES"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
 vCantonMascara = SIFGlobal.fxStringRelleno(vCantonMascara, "D", "0", rs!Caracteres)
End If
rs.Close

'Mascara del distrito
vDistritoMascara = "0"
strSQL = "select MAX(LEN(distrito)) as Caracteres from Distritos"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
 vDistritoMascara = SIFGlobal.fxStringRelleno(vDistritoMascara, "D", "0", rs!Caracteres)
End If
rs.Close


Call Formularios(Me)
Call RefrescaTags(Me)

Me.Icon = MDIPrincipal.Icon

Call sbRefrescaArbol

End Sub


Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

vModifica = "P"
vGrid.MaxRows = 0
vGrid.MaxCols = 2

lblNodeLinea(0).Tag = ""
lblNodeLinea(1).Tag = ""
lblNodeLinea(0).Caption = ""
lblNodeLinea(1).Caption = ""

lblY.Caption = "Provincias"


With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Provincias", "Provincias", "imgRoot")
  'Crear Arbol Inicial
  
    strSQL = "select * from provincias"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Call sbCreaNodos(vNode.Key, rs!Descripcion, "imgFolder", True, "N", "0x0" & rs!Provincia & "P")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With


End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim rsTmp As New ADODB.Recordset

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "Provincias" Then

Select Case Right(Node.Key, 1)
        
    Case "P" 'Provincias
    
        strSQL = "select * from cantones where Provincia = " & fxIndiceCodigo(Node.Key)
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          'Cantones
          Call sbCreaNodos("0x0" & rs!Provincia & "P", rs!Descripcion, "imgAsientos", False, "N", "0x0" & rs!Provincia & "-" & rs!Canton & "C")
          rs.MoveNext
        Loop
        rs.Close
    
    Case Else 'SubCuentas
     ''
End Select

End If

End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
               , vAcepta As String, Optional xkey As String = "N")
Dim nodx As Node, vKey As String
On Error Resume Next

Set nodx = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodx.Text = vTexto
    nodx.Tag = nodx.Index
    nodx.Image = vImagen
    If xkey = "N" Then
        nodx.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodx.Key = xkey
    End If
    
    
vKey = nodx.Key

If vExpand Then
    Set nodx = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodx.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodx.Tag = nodx.Index
End If
    
    
End Sub


Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim strSQL As String, i As Integer, vResulta As String
Dim vCadena As String, x As Integer

lblNodeLinea.Item(0).Tag = ""
lblNodeLinea.Item(1).Tag = ""

lblNodeLinea.Item(0).Caption = ""
lblNodeLinea.Item(1).Caption = ""

lblX.Caption = Node.FullPath
lblX.Tag = Node.Key


Select Case Right(Node.Key, 1)
   Case "P" 'Provincias - Carga Cantones
      vCadena = fxIndiceCodigo(Node.Key)
      strSQL = "select Canton,Descripcion from Cantones where Provincia = " & vCadena
   
      lblNodeLinea.Item(0).Tag = fxIndiceCodigo(Node.Key)
      lblNodeLinea.Item(0).Caption = "Provincia : " & Node.Text
      
      vModifica = "C"
      lblY.Caption = "Cantones"
   
   Case "C" 'Canton - Carga Distritos
      
      vModifica = "D"
      lblY.Caption = "Distritos"
      
      vCadena = fxIndiceCodigo(Node.Key)
      
      strSQL = "select Distrito,Descripcion from Distritos where Provincia = "
      
      
      vResulta = ""
      For i = 1 To Len(vCadena)
         If Mid(vCadena, i, 1) <> "-" Then
             vResulta = vResulta & Mid(vCadena, i, 1)
         Else
             Exit For
         End If
      Next i
      
      strSQL = strSQL & vResulta & " and Canton = '"
      
      
      lblNodeLinea.Item(0).Tag = vResulta
      lblNodeLinea.Item(0).Caption = "Provincia : " & Node.Parent
      
      
      vResulta = ""
      x = 0
      For i = 1 To Len(vCadena)
         If Mid(vCadena, i, 1) <> "-" Then
             If x = 1 Then vResulta = vResulta & Mid(vCadena, i, 1)
         Else
             x = x + 1
         End If
      Next i
      
      strSQL = strSQL & Format(vResulta, vCantonMascara) & "'"
      
      lblNodeLinea.Item(1).Tag = vResulta
      lblNodeLinea.Item(1).Caption = "Canton  : " & Node.Text
      
      
   Case Else
      'Carga Provincias es el Root
      strSQL = "select Provincia,Descripcion from Provincias"
      vModifica = "P"
      lblY.Caption = "Provincias"

End Select

Call sbCargaGrid(vGrid, 2, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow

vGrid.col = 1


Select Case vModifica
  Case "P"
        strSQL = "select isnull(count(*),0) as Existe from Provincias " _
               & " where Provincia = " & vGrid.Text
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then 'Insertar
          If Trim(vGrid.Text) = "" Then Exit Function
          
          strSQL = "insert into provincias(provincia,descripcion) values(" & vGrid.Text & ",'"
          vGrid.col = 2
          strSQL = strSQL & vGrid.Text & "')"
          Call ConectionExecute(strSQL)
        
          vGrid.col = 2
          Call Bitacora("Registra", "Provincia : " & Trim(vGrid.Text))
        
        Else 'Actualizar
        
         vGrid.col = 2
         strSQL = "update provincias set descripcion = '" & vGrid.Text & "' where Provincia = "
         vGrid.col = 1
         strSQL = strSQL & vGrid.Text
         Call ConectionExecute(strSQL)
         vGrid.col = 2
         Call Bitacora("Modifica", "Provincia : " & Trim(vGrid.Text))
        
        End If
        rs.Close
  
  Case "C"
        strSQL = "select isnull(count(*),0) as Existe from Cantones " _
               & " where Provincia = " & lblNodeLinea(0).Tag & " and Canton = '" & vGrid.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then 'Insertar
          If Trim(vGrid.Text) = "" Then Exit Function
          
          strSQL = "insert into cantones(provincia,canton,descripcion) values(" & lblNodeLinea(0).Tag & ",'" & vGrid.Text & "','"
          vGrid.col = 2
          strSQL = strSQL & vGrid.Text & "')"
          Call ConectionExecute(strSQL)
        
          vGrid.col = 2
          Call Bitacora("Registra", "Prov:" & lblNodeLinea(0).Tag & " Canton :" & Trim(vGrid.Text))
        
        Else 'Actualizar
        
         vGrid.col = 2
         strSQL = "update cantones set descripcion = '" & vGrid.Text & "' where Provincia = " & lblNodeLinea(0).Tag & " and Canton = '"
         vGrid.col = 1
         strSQL = strSQL & vGrid.Text & "'"
         Call ConectionExecute(strSQL)
         
         vGrid.col = 2
         Call Bitacora("Registra", "Prov:" & lblNodeLinea(0).Tag & " Canton :" & Trim(vGrid.Text))
        
        End If
        rs.Close
  
  
  Case "D"
        strSQL = "select isnull(count(*),0) as Existe from distritos " _
               & " where Provincia = " & lblNodeLinea(0).Tag & " and Canton = '" & Format(lblNodeLinea(1).Tag, vCantonMascara) _
               & "' and distrito = '" & Format(vGrid.Text, vDistritoMascara) & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then 'Insertar
          If Trim(vGrid.Text) = "" Then Exit Function
          
          strSQL = "insert into distritos(provincia,canton,distrito,descripcion) values(" & lblNodeLinea(0).Tag _
                 & ",'" & Format(lblNodeLinea(1).Tag, vCantonMascara) & "','" & Format(vGrid.Text, vDistritoMascara) & "','"
          vGrid.col = 2
          strSQL = strSQL & vGrid.Text & "')"
          Call ConectionExecute(strSQL)
        
          vGrid.col = 2
          Call Bitacora("Registra", "Prov:" & lblNodeLinea(0).Tag & "Cant:" & lblNodeLinea(1).Tag & " Dist:" & Trim(vGrid.Text))
        
        Else 'Actualizar
        
         vGrid.col = 2
         strSQL = "update distritos set descripcion = '" & vGrid.Text & "' where Provincia = " _
                & lblNodeLinea(0).Tag & " and Canton = '" & Format(lblNodeLinea(1).Tag, vCantonMascara) & "' and Distrito = '"
         vGrid.col = 1
         strSQL = strSQL & Format(vGrid.Text, vDistritoMascara) & "'"
         Call ConectionExecute(strSQL)
         
         vGrid.col = 2
         Call Bitacora("Registra", "Prov:" & lblNodeLinea(0).Tag & "Cant:" & lblNodeLinea(1).Tag & " Dist:" & Trim(vGrid.Text))
        
        End If
        rs.Close


End Select

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub imgRefrescar_Click()
Call sbRefrescaArbol
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

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

End Sub

