VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmBusquedaArticulos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Busqueda de Articulos"
   ClientHeight    =   6984
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10128
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6984
   ScaleWidth      =   10128
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4572
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   10092
      _Version        =   1245187
      _ExtentX        =   17801
      _ExtentY        =   8064
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      Sorted          =   -1  'True
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnOption 
      Height          =   372
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Código"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9240
      Top             =   960
   End
   Begin XtremeSuiteControls.FlatEdit txt 
      Height          =   312
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   10092
      _Version        =   1245187
      _ExtentX        =   17801
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnOption 
      Height          =   372
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cabys"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnOption 
      Height          =   372
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Descripción"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnOption 
      Height          =   372
      Index           =   3
      Left            =   6240
      TabIndex        =   6
      Top             =   1440
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Barras"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnOption 
      Height          =   372
      Index           =   4
      Left            =   8160
      TabIndex        =   7
      Top             =   1440
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Fabricante"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Articulos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   3852
   End
   Begin VB.Image imgBanner 
      Height          =   1236
      Left            =   0
      Top             =   0
      Width           =   10200
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
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select Top 50 cod_producto,cabys, cod_barras,cod_fabricante,descripcion,precio_regular,existencia,impuesto_ventas" _
       & " from pv_productos where estado = 'A' and "

Select Case True
  Case btnOption.Item(0).Checked
     strSQL = strSQL & " cod_producto like '%" & txt & "%' order by cod_producto"
  Case btnOption.Item(1).Checked
     strSQL = strSQL & " cabys like '%" & txt & "%' order by cabys"
  Case btnOption.Item(2).Checked
     strSQL = strSQL & " descripcion like '%" & txt & "%' order by descripcion"
  Case btnOption.Item(3).Checked
     strSQL = strSQL & " cod_barras like '%" & txt & "%' order by cod_barras"
  Case btnOption.Item(4).Checked
     strSQL = strSQL & " cod_fabricante like '%" & txt & "%' order by cod_fabricante"
End Select

lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_producto)
     itmX.SubItems(1) = rs!Cabys & ""
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = rs!cod_barras & ""
     itmX.SubItems(4) = Format(rs!precio_regular + (rs!precio_regular * rs!impuesto_ventas / 100), "Standard")
     itmX.SubItems(5) = rs!Existencia
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault

End Sub

Private Sub btnOption_Click(Index As Integer)
Dim i As Integer

For i = 0 To btnOption.Count - 1
    btnOption.Item(i).Checked = False
Next i

btnOption.Item(Index).Checked = True

Call sbBusqueda

End Sub

Private Sub Form_Load()
Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 2000, vbCenter
    .Add , , "Cabys", 2000, vbCenter
    .Add , , "Descripción", 4000
    .Add , , "Barras", 2000, vbCenter
    .Add , , "Precio", 1400, vbRightJustify
    .Add , , "Cantidad", 1100, vbCenter
End With

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error Resume Next

gBusquedas.Resultado = Item.Text
gBusquedas.Resultado2 = Item.SubItems(2)
Unload Me
End Sub

Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

txt.SetFocus
Call btnOption_Click(2)

End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbBusqueda
End If
End Sub



