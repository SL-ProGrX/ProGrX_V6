VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmUS_Explorer 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Explorador"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12705
   HelpContextID   =   1001
   Icon            =   "frmUSExplorer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   12705
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgBarra 
      Left            =   6120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":15162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":1B9C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":22226
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":28A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":2F2EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":35B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":3C3AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":42C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":49472
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":4FCD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   6720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":56536
            Key             =   "autorizado"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":5CD98
            Key             =   "restringido"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":635FA
            Key             =   "Opcion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":69E5C
            Key             =   "link"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":706BE
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":76F20
            Key             =   "user"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":7D782
            Key             =   "grupos"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":83FE4
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":8A846
            Key             =   "Clip"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":910A8
            Key             =   "UserDetalle"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":9790A
            Key             =   "GruposDetalle"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":9E16C
            Key             =   "OpcionesDetalle"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":A49CE
            Key             =   "formularios"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSExplorer.frx":AB230
            Key             =   "OpcionesNodos"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   582
      ButtonWidth     =   2249
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgBarra"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Nodo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refrescar"
            Key             =   "refrescar"
            Object.ToolTipText     =   "refreca la información del arbol"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reportes"
            Key             =   "reportes"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ayuda"
            Key             =   "ayuda"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalle"
            Key             =   "detalle"
            ImageIndex      =   5
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otorgados"
            Key             =   "permisos"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Permisos"
            Key             =   "Accesos"
            Object.ToolTipText     =   "Mantenimiento de Accesos"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Activos"
            Key             =   "Estado"
            Object.Tag             =   "'A'"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "activos"
                  Text            =   "Usuarios Activos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "inactivos"
                  Text            =   "Usuarios Inactivos"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "todos"
                  Text            =   "Todos los Usuarios"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Elimina"
                  Text            =   "Elimina Permisos de Inactivos"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   2160
      Left            =   5400
      ScaleHeight     =   940.557
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ListView lswExplorer 
      Height          =   2160
      Left            =   2640
      TabIndex        =   0
      Top             =   675
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   3810
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgExplorer"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   2160
      Left            =   0
      TabIndex        =   1
      Top             =   675
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   3810
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Tag             =   " TreeView:"
      Top             =   360
      Width           =   2610
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   2685
      TabIndex        =   4
      Tag             =   " ListView:"
      Top             =   360
      Width           =   3210
   End
   Begin VB.Image imgSplitter 
      Height          =   2145
      Left            =   2565
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
End
Attribute VB_Name = "frmUS_Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vDato As String
Dim mbMoving As Boolean
Const sglSplitLimit = 500

Private Function fxCodigoGrupo(strGrupo As String) As Integer
Dim rsX As New ADODB.Recordset, strSQL As String
strSQL = "select id_grupo from grupos where nombre = '" & strGrupo & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF Then
 fxCodigoGrupo = 0
Else
 fxCodigoGrupo = rsX!id_grupo
End If
rsX.Close
End Function

Private Function fxIndiceMultiple(xkey As String, vTipo As String) As String
Dim i As Long, strResultado As String, blnPaso As Boolean

xkey = fxIndiceCodigo(xkey)

blnPaso = True

If vTipo = "T" Then ' Tipo
  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xkey, i, 1)
    Else
     blnPaso = False
    End If
    i = i + 1
  Loop
  
Else 'Numero

  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) = "-" Then blnPaso = False
    i = i + 1
  Loop
  strResultado = Mid(xkey, i, 50) '50 es un default ningun asiento es tan largo

End If

fxIndiceMultiple = strResultado

End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim strOpciones As String

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

Select Case Node.Text
  Case "Grupos"
        rs.Open "select Nombre,ID_Grupo from grupos order by nombre", glogon.Conection, adOpenStatic
        Do While Not rs.EOF
         Call sbCreaNodos(Node.Key, rs!Nombre, "GruposDetalle", False, "0x0" & rs!id_grupo & "G")
         rs.MoveNext
        Loop
        rs.Close
 
  Case "Usuarios"
        strSQL = "select Nombre,UserID from usuarios where estado in(" _
               & tlbPrincipal.Buttons.Item(11).Tag & ") order by nombre"
            
        rs.Open strSQL, glogon.Conection, adOpenStatic
        
        Do While Not rs.EOF
         Call sbCreaNodos(Node.Key, UCase(rs!Nombre), "UserDetalle", False, "0x0" & rs!UserID & "U")
         rs.MoveNext
        Loop
        rs.Close
  
  Case "Opciones"
        rs.Open "select rtrim(Nombre) as nombre,Modulo from modulos order by nombre", glogon.Conection, adOpenStatic
        Do While Not rs.EOF
         Call sbCreaNodos(Node.Key, rs!Nombre, "OpcionesDetalle", True, "0x0" & rs!modulo & "M")
         rs.MoveNext
        Loop
        rs.Close
  Case Else
     
     Select Case Right(vNode.Key, 1)
        Case "M" 'Despliga Formularios
            rs.Source = "select * from formularios" _
                      & " where modulo = " & fxIndiceCodigo(vNode.Key) _
                      & " order by formulario"
            rs.Open , glogon.Conection, adOpenStatic
            Do While Not rs.EOF
             Call sbCreaNodos(Node.Key, rs!Descripcion, "formularios", True, "0x0" & rs!modulo & "-" & rs!frmID & "F")
             rs.MoveNext
            Loop
            rs.Close
            
        Case "F" 'Despliga Opciones
            rs.Open "select id_opt,Opcion,Opcion_descripcion as Descripcion" _
                  & " from Opciones where modulo = " & fxIndiceMultiple(vNode.Key, "T") _
                  & " and formulario in(select formulario from formularios where frmID = " _
                  & fxIndiceMultiple(Node.Key, "N") & ")  order by opcion" _
                  , glogon.Conection, adOpenStatic
            Do While Not rs.EOF

             Call sbCreaNodos(Node.Key, rs!Descripcion, "OpcionesNodos", False, "0x0" & rs!id_opt & "O")
             rs.MoveNext
            Loop
            rs.Close
     End Select

End Select

End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function


Private Sub sbMuestraDetalleSubNodos()
Dim itmX As ListItem, strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String

Select Case vNode.Parent
  Case "Grupos"
    
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
      
      lblTitle(1).Caption = lblTitle(1).Caption + "   - MIEMBROS"
      strSQL = "select nombre,fecha_miembro from miembros where id_grupo = " & fxIndiceCodigo(vNode.Key)   ' fxCodigoGrupo(vNode.Text)
      rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     With lswExplorer
      .ColumnHeaders.Add , , "Nombre", 4450
      .ColumnHeaders.Add , , "Fecha", 1450
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Nombre, , 4)
           itmX.SubItems(1) = Format(rs!fecha_miembro, "dd/mmm/yyyy")
       rs.MoveNext
      Loop
       rs.Close
     End With
      
    Else
    
      lblTitle(1).Caption = lblTitle(1).Caption + "   - PERMISOS"
              
              
      strSQL = "select G.nombre,O.*,M.nombre as ModuloNam,P.estado" _
             & " from Grupos G inner join Permisos P on G.id_grupo = P.nombre " _
             & " inner join Opciones O on P.id_opt= O.id_opt" _
             & " inner join Modulos M on O.modulo = M.modulo" _
             & " where P.tipo = 'G' and P.nombre = '" & fxIndiceCodigo(vNode.Key) & "'" _
             & " Order By M.Modulo,O.formulario"
      
     vCadena = ""
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     
     With lswExplorer
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Tipo", 1450
      Do While Not rs.EOF
       If vCadena <> Trim(rs!modulonam) Then
          vCadena = Trim(rs!modulonam)
        Set itmX = .ListItems.Add(, , rs!modulonam)
            itmX.ForeColor = vbBlue
            itmX.Bold = True
       End If
       
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!formulario, , 4)
           itmX.Tag = itmX.Index
           itmX.SubItems(1) = rs!opcion
           itmX.SubItems(2) = rs!Opcion_descripcion
           itmX.SubItems(3) = IIf((rs!Estado = "A"), "Autorización", "Restricción")
           If rs!Estado = "R" Then itmX.ForeColor = vbRed
       
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    End If
    
    
  Case "Usuarios"
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
      
     lblTitle(1).Caption = lblTitle(1).Caption + "   - MIEMBRO DE..."
     strSQL = " select M.*,G.nombre" _
            & " from miembros M inner join Grupos G on M.id_grupo = G.id_grupo" _
            & " where M.nombre in(select nombre from usuarios where userID = " _
            & fxIndiceCodigo(vNode.Key) & ")"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     With lswExplorer
      .ColumnHeaders.Add , , "Grupo", 4450
      .ColumnHeaders.Add , , "Fecha", 2450
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Nombre, , 9)
           itmX.SubItems(1) = Format(rs!fecha_miembro, "dddd, mmm d yyyy")
       rs.MoveNext
      Loop
       rs.Close
     End With
      
      
    Else
    
     strSQL = "Select O.formulario,O.opcion,O.opcion_descripcion,P.estado,M.nombre as ModuloNam,F.descripcion as FormX" _
            & " from Permisos P inner join Opciones O on P.id_opt = O.id_opt" _
            & " and P.nombre = '" & fxIndiceCodigo(vNode.Key) & "' and P.tipo = 'U'" _
            & " inner join Modulos M on O.modulo = M.modulo" _
            & " inner join formularios F on O.formulario = F.formulario" _
            & " group by O.formulario,O.opcion,O.opcion_descripcion,P.estado,M.nombre,F.descripcion"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     vCadena = ""
     With lswExplorer
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Tipo", 1450
      Do While Not rs.EOF
       
       If vCadena <> Trim(rs!modulonam) Then
          vCadena = Trim(rs!modulonam)
        Set itmX = .ListItems.Add(, , rs!modulonam)
            itmX.ForeColor = vbBlue
            itmX.Bold = True
       End If
       
       Set itmX = .ListItems.Add(, , rs!formx, , 9)
           itmX.SubItems(1) = rs!opcion
           itmX.SubItems(2) = rs!Opcion_descripcion
           itmX.SubItems(3) = IIf((rs!Estado = "A"), "Autorización", "Restricción")
           If rs!Estado = "R" Then itmX.ForeColor = vbRed
       
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    
    End If
  
  Case "Opciones"
 
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
  
      strSQL = "select * from opciones where modulo = " & fxIndiceCodigo(vNode.Key) & " order by formulario,opcion"

      rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     With lswExplorer
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!formulario, , 14)
           itmX.SubItems(1) = rs!opcion
           itmX.SubItems(2) = rs!Opcion_descripcion
       rs.MoveNext
      Loop
       rs.Close
     End With

    Else
    
      lblTitle(1).Caption = lblTitle(1).Caption + "   - PERMISOS OTORGADOS"

     strSQL = "O.*,G.Nombre as Grupo,P.estado" _
            & " from Permisos P inner join Opciones O on P.id_opt = O.id_opt" _
            & " inner join Grupos G on P.nombre = G.id_grupo and P.tipo = 'G'" _
            & " where O.modulo = " & fxIndiceCodigo(vNode.Key) & " order by O.formulario"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     With lswExplorer
      .ColumnHeaders.Add , , "Otrogado", 2800
      .ColumnHeaders.Add , , "Formulario", 3450
      .ColumnHeaders.Add , , "Opción", 2050
      .ColumnHeaders.Add , , "Descripción", 3450
      .ColumnHeaders.Add , , "Tipo", 1450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Grupo, , 7)
           itmX.SubItems(1) = rs!formulario
           itmX.SubItems(2) = rs!opcion
           itmX.SubItems(3) = rs!Descripcion
           itmX.SubItems(4) = IIf((rs!Estado = "A"), "Autorización", "Restricción")
           If rs!Estado = "R" Then itmX.ForeColor = vbRed
       rs.MoveNext
      Loop
       rs.Close
       
      strSQL = "select O.*,U.nombre as Usuario,P.estado" _
             & " from Permisos P inner join Opciones O on P.id_opt = O.id_opt" _
             & " inner join Usuarios U on P.UserID = U.userId and P.tipo = 'U'" _
             & " where O.modulo = " & fxIndiceCodigo(vNode.Key) & " order by O.formulario"
      rs.Open strSQL, glogon.Conection, adOpenForwardOnly
       
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Usuario, , 6)
           itmX.SubItems(1) = rs!formulario
           itmX.SubItems(2) = rs!opcion
           itmX.SubItems(3) = rs!Descripcion
           itmX.SubItems(4) = IIf((rs!Estado = "A"), "Autorización", "Restricción")
           If rs!Estado = "R" Then itmX.ForeColor = vbRed
       rs.MoveNext
      Loop
       rs.Close
       
       
     End With
    
    End If

End Select

End Sub

Private Sub sbMuestraDetalle()
Dim itmX As ListItem, strSQL As String, rs As New ADODB.Recordset
Dim strOpciones As String

lswExplorer.ListItems.Clear
lswExplorer.ColumnHeaders.Clear

rs.CursorLocation = adUseServer

Select Case vNode.Text
  
  Case "Grupos"
    
    If Me.tlbPrincipal.Buttons.Item(6).Value = tbrPressed Then
      strSQL = "Select * from grupos order by nombre"
      
      rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     
     With lswExplorer
      .ColumnHeaders.Add , , "ID"
      .ColumnHeaders.Add , , "Nombre", 4450
      .ColumnHeaders.Add , , "Fecha", 1450
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!id_grupo, , 7)
           itmX.SubItems(1) = rs!Nombre
           itmX.SubItems(2) = Format(rs!fecha_creacion, "dd/mm/yyyy")
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    Else
    
      strSQL = "select G.nombre,O.*,P.estado,F.descripcion as FormX" _
             & " from Grupos G inner join Permisos P on G.id_grupo = P.nombre" _
             & " inner join Opciones O on P.id_opt = O.id_opt" _
             & " inner join Formularios F on O.formulario = F.formulario" _
             & " where P.tipo = 'G' order by G.nombre"

     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     With lswExplorer
      .ColumnHeaders.Add , , "Grupo", 4450
      .ColumnHeaders.Add , , "Formulario", 4450
      .ColumnHeaders.Add , , "Opción", 2450
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Tipo", 1450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Nombre, , 7)
           itmX.SubItems(1) = rs!formx
           itmX.SubItems(2) = rs!opcion
           itmX.SubItems(3) = rs!Opcion_descripcion
           itmX.SubItems(4) = IIf((rs!Estado = "A"), "Autorización", "Restricción")
           If rs!Estado = "R" Then itmX.ForeColor = vbRed
       rs.MoveNext
      Loop
       rs.Close
     End With
    
    
    End If
    
  Case "Usuarios"
  
     strSQL = "select * from usuarios where estado in(" & Me.tlbPrincipal.Buttons.Item(11).Tag & ") order by nombre"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     With lswExplorer
      .ColumnHeaders.Add , , "Nombre", 2450
      .ColumnHeaders.Add , , "Estado", 1450
      .ColumnHeaders.Add , , "Descripción", 4450
      .ColumnHeaders.Add , , "Ingreso", 1450
      .ColumnHeaders.Add , , "Ult.Mov", 1450
      
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Nombre, , 6)
           itmX.SubItems(1) = IIf((rs!Estado = "A"), "Activo", "Inactivo")
           itmX.SubItems(2) = rs!Descripcion
           itmX.SubItems(3) = Format(rs!Fecha_Ingreso, "dd/mm/yyyy")
           itmX.SubItems(4) = Format(rs!Fecha_Mod, "dd/mm/yyyy")
       rs.MoveNext
      Loop
       rs.Close
     End With
      
  Case "Opciones"
     
     strSQL = "select * from modulos order by modulo"
     rs.Open strSQL, glogon.Conection, adOpenForwardOnly
     With lswExplorer
      .ColumnHeaders.Add , , "Modulo", 1450
      .ColumnHeaders.Add , , "Descripción", 4450
      Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!modulo, , 11)
           itmX.SubItems(1) = IIf(Len(Trim(rs!Descripcion)) = 0, rs!Nombre, rs!Descripcion)
       rs.MoveNext
      Loop
       rs.Close
     End With
 
 
  Case "US"
  
     With lswExplorer
      .ColumnHeaders.Add , , "Empresa", 2450
      .ColumnHeaders.Add , , "Servidor", 2450
      .ColumnHeaders.Add , , "B.D.", 2450
      .ColumnHeaders.Add , , "Usuario", 2450
      .ColumnHeaders.Add , , "Fecha", 1450
    
       Set itmX = .ListItems.Add(, , GLOBALES.gstrNombreEmpresa, , 1)
           itmX.SubItems(1) = glogon.Servidor
           itmX.SubItems(2) = glogon.BaseDatos
           itmX.SubItems(3) = glogon.Usuario
           itmX.SubItems(4) = Format(fxFechaServidor, "dd/mm/yyyy")
     End With
  
  Case Else
  
    Select Case Right(vNode.Key, 1)
      Case "M" 'Muestra Formularios
            strSQL = "select * from formularios" _
                   & " where modulo = " & fxIndiceCodigo(vNode.Key) _
                   & " order by formulario"
            rs.Open strSQL, glogon.Conection, adOpenForwardOnly
            With lswExplorer
             .ColumnHeaders.Add , , "Formulario", 2450
             .ColumnHeaders.Add , , "Descripción", 4450
             Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!formulario, , 13)
                  itmX.SubItems(1) = IIf(Len(Trim(rs!Descripcion)) = 0, rs!formulario, rs!Descripcion)
              rs.MoveNext
             Loop
              rs.Close
            End With
      
      Case "F" 'Muestra Opciones
            strSQL = "select id_opt,Opcion,Opcion_descripcion as Descripcion" _
                  & " from Opciones where modulo = " & fxIndiceMultiple(vNode.Key, "T") _
                  & " and formulario in(select formulario from formularios where frmID = " _
                  & fxIndiceMultiple(vNode.Key, "N") & ")  order by opcion"
            rs.Open strSQL, glogon.Conection, adOpenForwardOnly
            With lswExplorer
             .ColumnHeaders.Add , , "OpcionID", 1050
             .ColumnHeaders.Add , , "OpcionSys", 2450
             .ColumnHeaders.Add , , "Descripción", 4450
             Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!id_opt, , 14)
                  itmX.SubItems(1) = rs!opcion
                  itmX.SubItems(2) = rs!Descripcion
              rs.MoveNext
             Loop
              rs.Close
            End With
      
      
      Case "O" 'Muestra Asignacion
            With lswExplorer
             .ColumnHeaders.Add , , "", 5450
            
            Set itmX = .ListItems.Add(, , ">>> AUTORIZACIONES")
                itmX.Bold = True
                itmX.ForeColor = vbBlue
            Set itmX = .ListItems.Add(, , "")
            
            
            strSQL = "select G.nombre" _
                   & " from Grupos G inner join Permisos P on G.id_grupo = P.nombre" _
                   & " where P.tipo = 'G' and P.id_opt = " & fxIndiceCodigo(vNode.Key) _
                   & " and P.estado = 'A' group by G.nombre"
            rs.Open strSQL, glogon.Conection, adOpenStatic
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "GRUPOS", , 1)
                    itmX.Bold = True
                    itmX.ForeColor = vbBlue
            End If
             
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , Space(10) & rs!Nombre)
              rs.MoveNext
            Loop
             rs.Close
            
            strSQL = "select U.nombre" _
                  & " from Usuarios U inner join Permisos P on U.UserID = P.nombre" _
                  & " where P.tipo = 'U' and P.id_opt = " & fxIndiceCodigo(vNode.Key) _
                  & " and P.estado = 'A'"
            rs.Open strSQL, glogon.Conection, adOpenForwardOnly
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "")
                Set itmX = .ListItems.Add(, , "USUARIOS", , 1)
                    itmX.Bold = True
                    itmX.ForeColor = vbBlue
            End If
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , Space(10) & rs!Nombre)
              rs.MoveNext
            Loop
            rs.Close
            
            
            Set itmX = .ListItems.Add(, , "")
            Set itmX = .ListItems.Add(, , ">>> RESTRICCIONES")
                itmX.Bold = True
                itmX.ForeColor = vbRed
            Set itmX = .ListItems.Add(, , "")
            strSQL = "select G.nombre" _
                   & " from Grupos G inner join Permisos P on G.id_grupo = P.nombre" _
                   & " where P.tipo = 'G' and P.id_opt = " & fxIndiceCodigo(vNode.Key) _
                   & " and P.estado = 'R'"
            rs.Open strSQL, glogon.Conection, adOpenStatic
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "GRUPOS", , 2)
                    itmX.Bold = True
                    itmX.ForeColor = vbRed
            End If
             
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Nombre, , 6)
              rs.MoveNext
            Loop
             rs.Close
            
            strSQL = "select U.nombre" _
                  & " from Usuarios U inner join Permisos P on U.UserID = P.nombre" _
                  & " where P.tipo = 'U' and P.id_opt = " & fxIndiceCodigo(vNode.Key) _
                  & " and P.estado = 'R'"
            rs.Open strSQL, glogon.Conection, adOpenForwardOnly
            If Not rs.EOF And Not rs.BOF Then
                Set itmX = .ListItems.Add(, , "")
                Set itmX = .ListItems.Add(, , "USUARIOS", , 2)
                    itmX.Bold = True
                    itmX.ForeColor = vbRed
            End If
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Nombre, , 6)
              rs.MoveNext
            Loop
            rs.Close
            
            
            End With
      
      Case Else
        Call sbMuestraDetalleSubNodos
     End Select
End Select

End Sub

Private Sub ArbolExp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then
' Call PopupMenu(MDIMenu.mnuAcciones, , x, y)
'End If
End Sub

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)

On Error Resume Next


Set vNode = Node


lblTitle(0).Caption = UCase(vNode.Text)
lblTitle(1).Caption = vNode.FullPath

Call sbMuestraDetalle


End Sub


Private Sub sbReporteGrupos()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Seguridad"
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Reporte='Reporte Al  " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .ReportFileName = SIFGlobal.fxSIFPathReportes("SegListadoGrupos.rpt")
    .Connect = glogon.ConectRPT
    .PrintReport
End With
 
Me.MousePointer = vbDefault

End Sub


Private Sub sbReporteOpciones()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Seguridad"
 
     .Connect = glogon.ConectRPT
 
     .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(1) = "Reporte='Reporte Al  " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
          
          
     .Formulas(3) = "fxUsuario='USER :" & glogon.Usuario & "'"
     .Formulas(4) = "fxServidor='SERVER :" & glogon.Servidor & "'"
     .Formulas(5) = "fxBaseDatos='DATABASE :" & glogon.BaseDatos & "'"
          
          
If tlbPrincipal.Buttons(6).Value = tbrPressed Then
    'Detalle
        .ReportFileName = SIFGlobal.fxSIFPathReportes("SegListadoOpciones.rpt")
Else
    
    'Permisos
        .ReportFileName = SIFGlobal.fxSIFPathReportes("SegListadoOpcionesOtorgadas.rpt")
   
   
   Select Case Right(vNode.Key, 1)
     Case "M" 'Modulo
        .SelectionFormula = "{MODULOS.MODULO} = " & fxIndiceCodigo(vNode.Key)
     Case "F"
        .SelectionFormula = "{MODULOS.MODULO} = " & fxIndiceMultiple(vNode.Key, "T") _
                          & " and {FORMULARIOS.FRMID} = " & fxIndiceMultiple(vNode.Key, "N")
     Case "O"
        .SelectionFormula = "{OPCIONES.ID_OPT} = " & fxIndiceCodigo(vNode.Key)
   End Select
     
End If
     .PrintReport
End With
Me.MousePointer = vbDefault

End Sub



Public Sub sbButtonPopUp(i As Integer)

On Error GoTo vError

Select Case i
 Case 1 'Editar
   Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(1))
 Case 2 'Reportes
   Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(3))
 Case 3 'Permisos
   Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(9))
End Select

vError:

End Sub

Private Sub Form_Load()
vModulo = 13
 Call sbRefrescaArbol
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If UnloadMode = 0 Then
'      Cancel = True
'      Me.WindowState = 1
'   End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    
    'set the width
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    ArbolExp.Width = X
    imgSplitter.Left = X
    lswExplorer.Left = X + 40
    lswExplorer.Width = Me.Width - (ArbolExp.Width + 140)
    lblTitle(0).Width = ArbolExp.Width
    lblTitle(1).Left = lswExplorer.Left + 20
    lblTitle(1).Width = lswExplorer.Width - 40


    'set the top
    lswExplorer.Top = ArbolExp.Top
    imgSplitter.Top = ArbolExp.Top
    ArbolExp.Height = Me.Height - 1300

    lswExplorer.Height = ArbolExp.Height
    imgSplitter.Height = ArbolExp.Height
End Sub

Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
   
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub


Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim itmX As ListItem

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "US", "Root", "Root")
  vNode.Bold = True
  'Crear Arbol Inicial
  Call sbCreaNodos("US", "Grupos", "grupos", True)
  Call sbCreaNodos("US", "Usuarios", "user", True)
  Call sbCreaNodos("US", "Opciones", "Opcion", True)
  
  .Nodes(1).Expanded = True
  
     With lswExplorer
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "Empresa", 2450
      .ColumnHeaders.Add , , "Servidor", 2450
      .ColumnHeaders.Add , , "B.D.", 2450
      .ColumnHeaders.Add , , "Usuario", 2450
      .ColumnHeaders.Add , , "Fecha", 1450
       Set itmX = .ListItems.Add(, , GLOBALES.gstrNombreEmpresa, , 1)
           itmX.SubItems(1) = glogon.Servidor
           itmX.SubItems(2) = glogon.BaseDatos
           itmX.SubItems(3) = glogon.Usuario
           itmX.SubItems(4) = Format(fxFechaServidor, "dd/mm/yyyy")
     End With
  
End With

End Sub

Function fxIndice(Str As String) As String
Dim nodX As Node, lng As Long
On Error Resume Next
With ArbolExp
  For lng = 2 To .Nodes.Count
    Set nodX = .Nodes.Item("0x0" & lng)
    If nodX.Text = Str Then
     fxIndice = "0x0" & lng
     Exit Function
    End If
  Next lng
End With
fxIndice = "0"
End Function



Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
GLOBALES.gstrReporte = "ListadoUsuarios"
Select Case Button.Key
  Case "editar"
    Select Case vNode.Text
      Case "Grupos"
        frmUS_Grupos.Show
      
      Case "Usuarios"
        frmUS_Usuarios.Show
        
      Case "Opciones"
        frmUS_Opciones.Show
      
      Case Else
 
       If vNode.Index > 1 Then
           Select Case vNode.Parent
              Case "Grupos"
                    gEntidad.Tipo = "G"
                    gEntidad.Grupo = vNode.Text
                    gEntidad.GrpID = fxIndiceCodigo(vNode.Key)
                
                frmUS_Grupos.Show
              Case "Usuarios"
                    gEntidad.Tipo = "U"
                    gEntidad.Usuario = vNode.Text
                    gEntidad.UserID = fxIndiceCodigo(vNode.Key)
                frmUS_Usuarios.Show
              Case "Opciones"
                frmUS_Opciones.Show
            End Select
       End If
    End Select
    
  Case "Accesos"
       If vNode.Index > 1 Then
           Select Case vNode.Parent
              Case "Grupos"
                    gEntidad.Tipo = "G"
                    gEntidad.Grupo = vNode.Text
                    gEntidad.GrpID = fxIndiceCodigo(vNode.Key)
              
                    frmUS_DerechosNew.Show
                    frmUS_DerechosNew.cmdDeshacer_Click
              Case "Usuarios"
                    gEntidad.Tipo = "U"
                    gEntidad.Usuario = vNode.Text
                    gEntidad.UserID = fxIndiceCodigo(vNode.Key)
                    
                    frmUS_DerechosNew.Show
                    frmUS_DerechosNew.cmdDeshacer_Click
            End Select
       End If
    
  Case "refrescar"
    Call sbRefrescaArbol
    
    
  Case "reportes"
    
    Select Case vNode.Text
      Case "Grupos"
        Call sbReporteGrupos
      Case "Usuarios"
        frmUS_ReporteUsuarios.Show
      
      Case "Opciones"
        Call sbReporteOpciones
      Case Else
        If vNode.Index > 1 Then
            Select Case vNode.Parent
              Case "Grupos"
                Call sbReporteGrupos
              Case "Usuarios"
                frmUS_ReporteUsuarios.Show
                frmUS_ReporteUsuarios.txtUsuario = vNode.Text
              ' Case "Opciones"
              Case Else
                 Call sbReporteOpciones
              
            End Select
        End If
    End Select

   
   Case "detalle", "permisos"
      lblTitle(0).Caption = vNode.FullPath
      If vNode.Index > 1 Then
         lblTitle(1).Caption = UCase(vNode.Parent) & " : " & UCase(vNode.Text)
      Else
         lblTitle(1).Caption = vNode.Text
      End If
      Call sbMuestraDetalle

    Case "ayuda"
          frmContenedor.CD.HelpContext = Me.HelpContextID
          frmContenedor.CD.ShowHelp


End Select
End Sub

Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String

Select Case ButtonMenu.Key
 Case "activos"
    tlbPrincipal.Buttons.Item(11).Caption = "Activos"
    tlbPrincipal.Buttons.Item(11).Tag = "'A'"
    tlbPrincipal.Buttons.Item(11).Image = 8
    tlbPrincipal.Buttons.Item(11).ToolTipText = "Muestra solo los usuarios activos"
    
    Call sbRefrescaArbol
    
 Case "inactivos"
    tlbPrincipal.Buttons.Item(11).Caption = "Inactivos"
    tlbPrincipal.Buttons.Item(11).Tag = "'I'"
    tlbPrincipal.Buttons.Item(11).Image = 9
    tlbPrincipal.Buttons.Item(11).ToolTipText = "Muestra solo los usuarios Inactivos"
    
    Call sbRefrescaArbol
    
 Case "todos"
    tlbPrincipal.Buttons.Item(11).Caption = "Inactivos"
    tlbPrincipal.Buttons.Item(11).Tag = "'A','I'"
    tlbPrincipal.Buttons.Item(11).Image = 10
    tlbPrincipal.Buttons.Item(11).ToolTipText = "Muestra TODOS los usuarios"
    
    Call sbRefrescaArbol
    
 Case "Elimina"
   Me.MousePointer = vbHourglass
   
   strSQL = "delete permisos" _
          & " where tipo = 'U' and nombre in(select UserId from usuarios where estado = 'I')"
   glogon.Conection.Execute strSQL
   
   strSQL = "delete miembros" _
          & " where nombre in(select Nombre from usuarios where estado = 'I')"
   glogon.Conection.Execute strSQL
   
    MsgBox "Actualización realizada satisfactoriamente...", vbInformation
   
   Me.MousePointer = vbDefault
 

End Select


End Sub
