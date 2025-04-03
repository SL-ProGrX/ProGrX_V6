VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPGX_DistribucionPolitica 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribución Política"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12255
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   5880
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   5040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":11B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":14D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":17F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":1B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":1E28
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":2144
            Key             =   "imgUsuario"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":2A20
            Key             =   "imgGrupo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":32FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":3618
            Key             =   "imgCuentas"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":3934
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":3A41
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":3B61
            Key             =   "imgFolder"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPGX_DistribucionPolitica.frx":3C7D
            Key             =   "imgAsientos"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5895
      Left            =   5400
      TabIndex        =   6
      Top             =   600
      Width           =   6735
      _Version        =   524288
      _ExtentX        =   11880
      _ExtentY        =   10398
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
      SpreadDesigner  =   "frmPGX_DistribucionPolitica.frx":3D8B
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Image imgRefrescar 
      Height          =   255
      Left            =   11760
      Picture         =   "frmPGX_DistribucionPolitica.frx":4318
      Stretch         =   -1  'True
      Top             =   240
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   5640
      X2              =   12000
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   240
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblNodeLinea 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Linea"
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label lblNodeLinea 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Linea"
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Distribución Política"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   ">> <<"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10080
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPGX_DistribucionPolitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vModifica As String
Dim vNodeCodex(4) As String, vCantonMascara As String, vDistritoMascara As String

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 13
vGrid.AppearanceStyle = fxGridStyle

vCantonMascara = ""
vDistritoMascara = ""

Call Formularios(Me)
Call RefrescaTags(Me)


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
  Set vNode = .Nodes.Add(, , "País", "País", "imgRoot")
  'Crear Arbol Inicial
  
    strSQL = "select * from PGX_PAIS"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Call sbCreaNodos(vNode.Key, rs!Descripcion, "imgFolder", True, "N", "0x0" & rs!cod_Pais & "P")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With


End Sub


Private Function fxIndiceCodigo(xKey As String) As String
xKey = Mid(xKey, 4, Len(xKey))
xKey = Mid(xKey, 1, Len(xKey) - 1)
fxIndiceCodigo = xKey
End Function

Private Function fxIndiceMultiple(xKey As String, pNivel As Integer) As String
Dim i As Long, strResultado As String, blnPaso As Boolean
Dim pLLaveCount As Integer

xKey = fxIndiceCodigo(xKey)
pLLaveCount = 0

blnPaso = True
i = 1
strResultado = ""

Do While blnPaso
  If Mid(xKey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xKey, i, 1)
  Else
    pLLaveCount = pLLaveCount + 1
    If pLLaveCount = pNivel Then
       blnPaso = False
    Else
       strResultado = "" 'Inicializa para el Nivel Siguiente
    End If
  End If
  i = i + 1
Loop
  

fxIndiceMultiple = strResultado

End Function

Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim rsTmp As New ADODB.Recordset

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "País" Then

Select Case Right(Node.Key, 1)
        
    Case "P" 'Pais: Expande N1
        strSQL = "select Descripcion, rTrim(Cod_Pais) + 'P' as 'NodoPadre'" _
               & ", rtrim(Cod_Pais) + '-' + rtrim(Cod_Pais_N1) + '1'  'LlaveNodo'" _
               & " from PGX_PAIS_N1 where cod_Pais = '" & fxIndiceCodigo(Node.Key) & "'"
               
    Case "1" 'Expande N2
        strSQL = "select Descripcion, rtrim(Cod_Pais) + '-' + rtrim(Cod_Pais_N1) + '1' as 'NodoPadre'" _
               & ", rtrim(Cod_Pais) + '-' + rtrim(Cod_Pais_N1) + rtrim(Cod_Pais_N2)  + '2'  'LlaveNodo'" _
               & " from PGX_PAIS_N2 where cod_Pais = '" & fxIndiceCodigo(Node.Key) & "'" _
               & " and cod_Pais_N1 = '" & fxIndiceMultiple(Node.Key, 1) & "'"

    Case "2" 'Expande N3
        strSQL = "select Descripcion, rtrim(Cod_Pais) + '-' + rtrim(Cod_Pais_N1) + '-' + rtrim(Cod_Pais_N2) + '2' as 'NodoPadre'" _
               & ", rtrim(Cod_Pais) + '-' + rtrim(Cod_Pais_N1) + " - " +  rtrim(Cod_Pais_N2) + " - " + rtrim(Cod_Pais_N3) + '3'  'LlaveNodo'" _
               & " from PGX_PAIS_N3 where cod_Pais = '" & fxIndiceCodigo(Node.Key) & "'" _
               & " and cod_Pais_N1 = '" & fxIndiceMultiple(Node.Key, 1) & "'" _
               & " and cod_Pais_N2 = '" & fxIndiceMultiple(Node.Key, 2) & "'"
    Case Else 'SubCuentas
     ''
End Select


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  'Cantones
  Call sbCreaNodos("0x0" & rs!NodoPadre, rs!Descripcion, "imgAsientos", False, "N", "0x0" & rs!LLaveNodo)
  rs.MoveNext
Loop
rs.Close


End If

End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
               , vAcepta As String, Optional xKey As String = "N")
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xKey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xKey
    End If
    
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
    
End Sub


Public Function fxPGX_Pais_Nivel(pPais As String, pNivel As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset

If pNivel = 0 Then
  strSQL = "País"
End If

If pNivel > 0 Then
    strSQL = "select N" & pNivel & "_Nombre as 'Nombre' from PGX_Pais where cod_Pais = '" & pPais & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If rs.EOF And rs.BOF Then
      strSQL = ""
    Else
      strSQL = Trim(rs!Nombre)
    End If
End If

fxPGX_Pais_Nivel = strSQL

End Function

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim strSQL As String, i As Integer, vResulta As String
Dim vCadena As String, x As Integer

lblNodeLinea.Item(0).Tag = ""
lblNodeLinea.Item(1).Tag = ""

lblNodeLinea.Item(0).Caption = ""
lblNodeLinea.Item(1).Caption = ""

lblX.Caption = Node.FullPath
lblX.Tag = Node.Key

vNodeCodex(0) = ""
vNodeCodex(1) = ""
vNodeCodex(2) = ""
vNodeCodex(3) = ""
vNodeCodex(4) = "1"

Select Case Right(Node.Key, 1)
   Case "P" 'Paises: Depliega Nivel 1
   
      vNodeCodex(0) = fxIndiceCodigo(Node.Key)
      vNodeCodex(4) = "1"
      
      strSQL = "select Cod_Pais_N1,Descripcion,Activo from PGX_PAIS_N1 where Cod_Pais = '" & vNodeCodex(0) & "'"
   
      lblNodeLinea.Item(0).Tag = vNodeCodex(0)
      lblNodeLinea.Item(0).Caption = "País : " & Node.Text
      
      lblY.Caption = fxPGX_Pais_Nivel(vNodeCodex(0), 0)
   
   Case "1" 'Nivel 1: Cantones
      
      vNodeCodex(0) = fxIndiceCodigo(Node.Key)
      vNodeCodex(1) = fxIndiceMultiple(Node.Key, 1)
      vNodeCodex(4) = "2"
      
      
      strSQL = "select Cod_Pais_N2,Descripcion,Activo " _
             & " from PGX_PAIS_N2 where Cod_Pais = '" & vNodeCodex(0) _
             & "' and cod_Pais_N1 = '" & vNodeCodex(1) & "'"
      
      
      lblY.Caption = fxPGX_Pais_Nivel(vNodeCodex(0), 1)
      lblNodeLinea.Item(0).Tag = vNodeCodex(1)
      lblNodeLinea.Item(0).Caption = fxPGX_Pais_Nivel(vNodeCodex(0), 1) & " : " & Node.Parent

   Case "2" 'Nivel 2: Distritos
      
      vNodeCodex(0) = fxIndiceCodigo(Node.Key)
      vNodeCodex(1) = fxIndiceMultiple(Node.Key, 1)
      vNodeCodex(2) = fxIndiceMultiple(Node.Key, 2)
      vNodeCodex(4) = "3"
      
      
      strSQL = "select Cod_Pais_N2,Descripcion,Activo " _
             & " from PGX_PAIS_N3 where Cod_Pais = '" & vNodeCodex(0) _
             & "' and cod_Pais_N1 = '" & vNodeCodex(1) & "'" _
             & "' and cod_Pais_N2 = '" & vNodeCodex(2) & "'"
      
      
      lblY.Caption = fxPGX_Pais_Nivel(vNodeCodex(0), 2)
      lblNodeLinea.Item(0).Tag = vNodeCodex(2)
      lblNodeLinea.Item(0).Caption = fxPGX_Pais_Nivel(vNodeCodex(0), 2) & " : " & Node.Parent



End Select


If vNodeCodex(0) <> "" Then
    Call sbCargaGrid(vGrid, 3, strSQL)
End If

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow

vGrid.Col = 1


Select Case vModifica
  Case "P"
        strSQL = "select isnull(count(*),0) as Existe from Provincias " _
               & " where Provincia = " & vGrid.Text
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then 'Insertar
          If Trim(vGrid.Text) = "" Then Exit Function
          
          strSQL = "insert into provincias(provincia,descripcion) values(" & vGrid.Text & ",'"
          vGrid.Col = 2
          strSQL = strSQL & vGrid.Text & "')"
          Call ConectionExecute(strSQL)
        
          vGrid.Col = 2
          Call Bitacora("Registra", "Provincia : " & Trim(vGrid.Text))
        
        Else 'Actualizar
        
         vGrid.Col = 2
         strSQL = "update provincias set descripcion = '" & vGrid.Text & "' where Provincia = "
         vGrid.Col = 1
         strSQL = strSQL & vGrid.Text
         Call ConectionExecute(strSQL)
         vGrid.Col = 2
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
          vGrid.Col = 2
          strSQL = strSQL & vGrid.Text & "')"
          Call ConectionExecute(strSQL)
        
          vGrid.Col = 2
          Call Bitacora("Registra", "Prov:" & lblNodeLinea(0).Tag & " Canton :" & Trim(vGrid.Text))
        
        Else 'Actualizar
        
         vGrid.Col = 2
         strSQL = "update cantones set descripcion = '" & vGrid.Text & "' where Provincia = " & lblNodeLinea(0).Tag & " and Canton = '"
         vGrid.Col = 1
         strSQL = strSQL & vGrid.Text & "'"
         Call ConectionExecute(strSQL)
         
         vGrid.Col = 2
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
          vGrid.Col = 2
          strSQL = strSQL & vGrid.Text & "')"
          Call ConectionExecute(strSQL)
        
          vGrid.Col = 2
          Call Bitacora("Registra", "Prov:" & lblNodeLinea(0).Tag & "Cant:" & lblNodeLinea(1).Tag & " Dist:" & Trim(vGrid.Text))
        
        Else 'Actualizar
        
         vGrid.Col = 2
         strSQL = "update distritos set descripcion = '" & vGrid.Text & "' where Provincia = " _
                & lblNodeLinea(0).Tag & " and Canton = '" & Format(lblNodeLinea(1).Tag, vCantonMascara) & "' and Distrito = '"
         vGrid.Col = 1
         strSQL = strSQL & Format(vGrid.Text, vDistritoMascara) & "'"
         Call ConectionExecute(strSQL)
         
         vGrid.Col = 2
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

