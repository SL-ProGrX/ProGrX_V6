VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusquedaArticulos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Busqueda de Articulos"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBusquedaArticulos.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   8055
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      Caption         =   "Código de Barras"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      Caption         =   "Código del Fabricante"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   14737632
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Barras"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripción"
         Object.Width           =   6950
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Precio"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Cantidad"
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   9000
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Consulta de Articulos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar Articulos por "
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
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmBusquedaArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbBusqueda()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select cod_producto,cod_barras,cod_fabricante,descripcion,precio_regular,existencia,impuesto_ventas" _
       & " from pv_productos where estado = 'A' and "

Select Case True
  Case opt.Item(0)
     strSQL = strSQL & " cod_producto like '%" & txt & "%' order by cod_producto"
  Case opt.Item(1)
     strSQL = strSQL & " descripcion like '%" & txt & "%' order by descripcion"
  Case opt.Item(2)
     strSQL = strSQL & " cod_barras like '%" & txt & "%' order by cod_barras"
  Case opt.Item(3)
     strSQL = strSQL & " cod_fabricante like '%" & txt & "%' order by cod_fabricante"
End Select

lsw.ListItems.Clear
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 DoEvents
 Set itmX = lsw.ListItems.Add(, , rs!cod_producto)
     itmX.SubItems(1) = rs!cod_barras
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = Format(rs!precio_regular + (rs!precio_regular * rs!impuesto_ventas / 100), "Standard")
     itmX.SubItems(4) = rs!Existencia
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""

End Sub

Private Sub lsw_Click()
On Error Resume Next

gBusquedas.Resultado = lsw.SelectedItem.Text
gBusquedas.Resultado2 = lsw.SelectedItem.SubItems(2)
Unload Me

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo Errores
    lsw.SortKey = ColumnHeader.Index - 1
    
    If (lsw.SortOrder = lvwAscending) Then
        lsw.SortOrder = lvwDescending
    Else
        lsw.SortOrder = lvwAscending
    End If
    
    lsw.Sorted = True
Salir:
    Exit Sub
Errores:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbBusqueda
End If
End Sub



