VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmFNDTrasladoFondos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslados de Fondos"
   ClientHeight    =   4944
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9252
   Icon            =   "frmFNDTrasladoFondos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4944
   ScaleWidth      =   9252
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   3732
      Left            =   -84
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   9372
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3012
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   7572
         _Version        =   1245187
         _ExtentX        =   13356
         _ExtentY        =   5313
         _StockProps     =   77
         BackColor       =   -2147483643
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
         ShowBorder      =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   7572
         _Version        =   1245187
         _ExtentX        =   13356
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Planes disponibles para Traslado de Fondos:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnOrigen 
      Height          =   312
      Left            =   480
      TabIndex        =   17
      Top             =   1680
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Origen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox gbAplicar 
      Height          =   972
      Left            =   480
      TabIndex        =   9
      Top             =   3960
      Width           =   8412
      _Version        =   1245187
      _ExtentX        =   14838
      _ExtentY        =   1714
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   612
         Left            =   6960
         TabIndex        =   10
         Top             =   240
         Width           =   1452
         _Version        =   1245187
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "&Aplicar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmFNDTrasladoFondos.frx":000C
      End
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   7332
      _Version        =   1245187
      _ExtentX        =   12933
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtOContrato 
      Height          =   312
      Left            =   1560
      TabIndex        =   11
      Top             =   1680
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOPlan 
      Height          =   312
      Left            =   2880
      TabIndex        =   12
      Top             =   1680
      Width           =   6012
      _Version        =   1245187
      _ExtentX        =   10604
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDContrato 
      Height          =   312
      Left            =   1560
      TabIndex        =   13
      Top             =   2040
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDPlan 
      Height          =   312
      Left            =   2880
      TabIndex        =   14
      Top             =   2040
      Width           =   6012
      _Version        =   1245187
      _ExtentX        =   10604
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1560
      TabIndex        =   15
      Top             =   240
      Width           =   1692
      _Version        =   1245187
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3240
      TabIndex        =   16
      Top             =   240
      Width           =   5652
      _Version        =   1245187
      _ExtentX        =   9970
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnDestino 
      Height          =   312
      Left            =   480
      TabIndex        =   18
      Top             =   2040
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Destino"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   792
      Left            =   1560
      TabIndex        =   19
      Top             =   2400
      Width           =   7332
      _Version        =   1245187
      _ExtentX        =   12933
      _ExtentY        =   1397
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDisponible 
      Height          =   312
      Left            =   3240
      TabIndex        =   20
      Top             =   3360
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   6960
      TabIndex        =   21
      Top             =   3360
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   2
      Left            =   2880
      TabIndex        =   23
      Top             =   1320
      Width           =   6012
      _Version        =   1245187
      _ExtentX        =   10604
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Descripción"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   1
      Left            =   1560
      TabIndex        =   22
      Top             =   1320
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "No. Contrato"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   6
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a Trasladar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   5
      Left            =   5040
      TabIndex        =   4
      Top             =   3360
      Width           =   1932
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   4
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Tag             =   "Op"
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
End
Attribute VB_Name = "frmFNDTrasladoFondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vTipo As String

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub


Private Sub btnDestino_Click()
vTipo = "D"
Call sbLlenaLsw
End Sub

Private Sub btnOrigen_Click()
vTipo = "O"
Call sbLlenaLsw
End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Dim vTipoDoc  As String, vNumDoc    As Long

On Error GoTo vError

MousePointer = vbHourglass

strSQL = "exec spFndTrasladosFondos '" & Trim(txtOPlan.Tag) & "'," & Trim(txtOContrato) & ",'" & Trim(txtCedula.Text) _
        & "'," & CCur(txtMonto.Text) & ",'" & txtDPlan.Tag & "'," & Trim(txtDContrato.Text) & ",'" & Trim(txtCedula.Text) _
        & "','" & glogon.Usuario & "','" & txtNotas.Text & "','ProGrX'"
Call OpenRecordSet(rs, strSQL)
    vTipoDoc = rs!TipoDoc
    vNumDoc = rs!NumDoc
rs.Close

MousePointer = vbDefault

'Imprime Documentos
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

MsgBox "Traslado Aplicado, Nota de Traslado Procesada # " & vNumDoc, vbInformation
 
 Call sbLimpiaDatos

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

Call sbLimpiaDatos

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbLimpiaDatos()

fra.Visible = False
txtOContrato = 0
txtOPlan = ""
txtOPlan.Tag = ""

txtDContrato = 0
txtDPlan = ""
txtDPlan.Tag = ""

txtDisponible = 0
txtMonto = 0

txtNotas = ""

End Sub

Private Sub sbLlenaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

fra.Visible = True
fra.Left = 0

lsw.ListItems.Clear
With lsw.ColumnHeaders
    .Add , , "Contrato No", 1400
    .Add , , "Plan", 1400, vbCenter
    .Add , , "Descripción", 3200
    .Add , , "Disponible", 1500, vbRightJustify
End With

strSQL = "select C.cod_contrato,C.cod_plan,C.aportes + C.rendimiento - isnull(C.Monto_Transito,0) as 'Disponible' ,P.descripcion" _
       & " from fnd_planes P inner join fnd_contratos C on P.cod_plan = C.cod_plan" _
       & " and P.cod_operadora = C.cod_operadora" _
       & " where C.estado = 'A' and C.cod_operadora = " & cbo.ItemData(cbo.ListIndex) _
       & " and C.cedula = '" & txtCedula & "' and isnull(P.MOV_ENTRE_FONDOS,0) = 1"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!COD_CONTRATO)
     itmX.SubItems(1) = rs!cod_Plan
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = Format(rs!Disponible, "Standard")
 rs.MoveNext
Loop
rs.Close

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError

Select Case vTipo
 Case "O"
    txtOContrato = Item.Text
    txtOPlan = Item.SubItems(2)
    txtOPlan.Tag = Item.SubItems(1)
    txtDisponible = Item.SubItems(3)
 Case "D"
    txtDContrato = Item.Text
    txtDPlan = Item.SubItems(2)
    txtDPlan.Tag = Item.SubItems(1)
End Select

fra.Visible = False

vError:
End Sub


Private Sub txtCedula_Change()
Call sbLimpiaDatos
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto = CCur(txtMonto)

vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto = Format(CCur(txtMonto), "Standard")

vError:

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub

