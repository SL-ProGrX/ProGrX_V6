VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_NivelesResolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Niveles de Resolución"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   HelpContextID   =   3021
   Icon            =   "frmCR_NivelesResolucion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   6135
   Begin VB.Timer tmrMensajes 
      Interval        =   300
      Left            =   5640
      Top             =   5280
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5520
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_NivelesResolucion.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_NivelesResolucion.frx":0BEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbGuardar 
      Height          =   570
      Left            =   5520
      TabIndex        =   8
      Top             =   0
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Asigna este Nivel de Autorización"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lswUsuarios 
      Height          =   2175
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuario"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.TextBox txtHasta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtDesde 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtUsuario 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_NivelesResolucion.frx":0F12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lswNiveles 
      Height          =   2535
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuario"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Desde"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Hasta"
         Object.Width           =   2893
      EndProperty
   End
   Begin VB.Label lblMensajes 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5760
      Width           =   6135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USUARIOS CON RANGOS DE RESOLUCION ASIGNADOS"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LISTA DE USUARIOS SIN ASIGNAR"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HASTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESDE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmCR_NivelesResolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMensaje As String

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
vModulo = 3
Call Formularios(Me)

strMensaje = "[Para Borrar un Usuario de los Niveles de Resolución" _
           & " debe de poner el mismo valor en los rangos DESDE y HASTA ]   "
Me.MousePointer = 11
 CargaNiveles
 CargaUsuarios
Me.MousePointer = 1
Call RefrescaTags(Me)
End Sub

Sub CargaUsuarios()
Dim rs As New ADODB.Recordset, itmX As ListItem, i As Long

rs.Open "select * from usuarios", glogon.Conection, adOpenStatic

lswUsuarios.ListItems.Clear

i = 1
Do While rs.EOF = False
 If fxNoExisteNivel(rs!Nombre) Then
   Set itmX = lswUsuarios.ListItems.Add(i, , rs!Nombre, , 1)
       itmX.SubItems(1) = rs!Descripcion
    i = i + 1
 End If
 rs.MoveNext
Loop
rs.Close

End Sub

Sub CargaNiveles()
Dim rs As New ADODB.Recordset, itmX As ListItem

rs.CursorLocation = adUseServer
rs.Open "select * from niveles_resolutivos", glogon.Conection, adOpenStatic

lswNiveles.ListItems.Clear

Do While rs.EOF = False
 Set itmX = lswNiveles.ListItems.Add(lswNiveles.ListItems.Count + 1, , rs!Nombre, , 2)
     itmX.SubItems(1) = Format(rs!desde, "###,###,###,##0.00")
     itmX.SubItems(2) = Format(rs!hasta, "###,###,###,##0.00")
     itmX.Tag = itmX.Index
 rs.MoveNext
Loop

rs.Close

End Sub

Sub LimpiaDatos()

txtUsuario = ""
txtDesde = ""
txtHasta = ""

End Sub


Private Sub lswNiveles_Click()
On Error Resume Next
With lswNiveles
 txtUsuario = .SelectedItem.Text
 txtDesde = CCur(.SelectedItem.SubItems(1))
 txtHasta = CCur(.SelectedItem.SubItems(2))
End With
 txtDesde.SetFocus
End Sub

Private Sub lswNiveles_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 Set Conlsw.lswX = lswNiveles
 Conlsw.Abre
End If
End Sub

Private Sub lswUsuarios_Click()
On Error Resume Next
With lswUsuarios
 txtUsuario = .SelectedItem.Text
 txtDesde = 0
 txtHasta = 0
End With
 txtDesde.SetFocus
End Sub

Private Sub tlbGuardar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String
Dim iBitacora As Integer
'select case but
If fxValidacion() Then

Me.MousePointer = 11
iBitacora = 0

If txtDesde <> txtHasta Then
    If fxNoExisteNivel(txtUsuario) Then
     iBitacora = 1
     strSQL = "insert into niveles_resolutivos(nombre,desde,hasta) values" _
         & "('" & Trim(txtUsuario) & "'," & CCur(txtDesde) & "," & CCur(txtHasta) & ")"
     
    Else
     iBitacora = 2
     strSQL = "update niveles_resolutivos set desde = " & CCur(txtDesde) _
         & ", hasta = " & CCur(txtHasta) & " where nombre = '" & Trim(txtUsuario) & "'"
    End If
Else
    iBitacora = 3
    strSQL = "delete niveles_resolutivos where nombre = '" & Trim(txtUsuario) & "'"
End If ' <>

glogon.Conection.Execute strSQL

On Error Resume Next

If iBitacora = 1 Then
    Call Bitacora("Registra", "Nivel Resolutivo D:" & CCur(txtDesde) & " H:" & CCur(txtHasta) & " US:" & Trim(txtUsuario))
ElseIf iBitacora = 2 Then
    Call Bitacora("Modifica", "Nivel Resolutivo D:" & CCur(txtDesde) & " H:" & CCur(txtHasta) & " US:" & Trim(txtUsuario))
ElseIf iBitacora = 3 Then
    Call Bitacora("Borra", "Nivel Resolutivo US:" & Trim(txtUsuario))
End If

CargaNiveles
CargaUsuarios
LimpiaDatos

Me.MousePointer = 1

Else
 
 MsgBox "Valores Desde y Hasta no son Válidos", vbCritical

End If 'validacion
Call RefrescaTags(Me)
End Sub


Function fxNoExisteNivel(str As String) As Boolean
Dim rsFx As New ADODB.Recordset
str = Trim(str)

rsFx.Open "select count(*) as Existe from niveles_resolutivos where nombre = '" & str & "'", glogon.Conection, adOpenStatic
If rsFx!existe = 0 Then
 fxNoExisteNivel = True
Else
 fxNoExisteNivel = False
End If
rsFx.Close

End Function

Function fxValidacion() As Boolean
Dim curDesde As Currency, curHasta As Currency

If Trim(txtUsuario) <> "" Then

curDesde = CCur(txtDesde)
curHasta = CCur(txtHasta)

If curDesde > curHasta Then
 fxValidacion = False
 Exit Function
End If

If curDesde < 0 Then
 fxValidacion = False
 Exit Function
End If

If curHasta < 0 Then
 fxValidacion = False
 Exit Function
End If

If Trim(txtUsuario) = "" Then
 fxValidacion = False
 Exit Function
End If

Else
  fxValidacion = False
  Exit Function
End If

fxValidacion = True

End Function


Private Sub tmrMensajes_Timer()
 strMensaje = Mid(strMensaje, 2, Len(strMensaje)) + Mid(strMensaje, 1, 1)
 
 lblMensajes.Caption = strMensaje
 lblMensajes.Refresh
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtHasta.SetFocus
End Sub

