VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmAF_BeneProdPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega de Productos Beneficios Asignados"
   ClientHeight    =   7380
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10620
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10620
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5532
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   10332
      _Version        =   1245187
      _ExtentX        =   18224
      _ExtentY        =   9758
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox frDetalle 
      Height          =   4572
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   9852
      _Version        =   1245187
      _ExtentX        =   17378
      _ExtentY        =   8064
      _StockProps     =   79
      Caption         =   "...."
      ForeColor       =   8421504
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lswDetalle 
         Height          =   3492
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   9612
         _Version        =   1245187
         _ExtentX        =   16954
         _ExtentY        =   6159
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
      End
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   372
         Left            =   8400
         TabIndex        =   11
         Top             =   3960
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
   End
   Begin VB.CheckBox chkDetalle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "&Detalle de Productos"
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
      Height          =   252
      Left            =   8160
      TabIndex        =   3
      Top             =   1320
      Width           =   2172
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "&Todos"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   156
      Left            =   0
      TabIndex        =   0
      Top             =   7224
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdGenerar 
      Height          =   372
      Left            =   8880
      TabIndex        =   5
      Top             =   480
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Generar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   372
      Left            =   7560
      TabIndex        =   6
      Top             =   480
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   5172
      _Version        =   1245187
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Beneficio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Seleccione los Beneficios (Productos)"
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
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   10332
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_BeneProdPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Dim i As Integer

Private Sub cbo_Click()
lsw.ListItems.Clear

End Sub

Private Sub chkTodos_Click()

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

lsw.ListItems.Clear

chkTodos.Value = vbUnchecked
        
                
strSQL = "Select A.consec,A.cod_beneficio,  O.cedula, isnull(S.nombre,'') as 'Nombre'" _
       & " ,sum(A.cantidad) as cantidad,sum(A.costo_unidad*A.cantidad) as 'Monto'" _
       & " from afi_bene_prodasg A" _
       & " left join afi_bene_otorga O on A.cod_beneficio = O.cod_Beneficio and A.Consec = O.Consec" _
       & " left join Socios S on O.cedula = S.cedula" _
       & " where O.estado = 'S' and A.cod_beneficio = '" & cbo.ItemData(cbo.ListIndex) & "'" _
       & " group by A.consec,A.cod_beneficio, O.Cedula, S.nombre"
        
Call OpenRecordSet(rs, strSQL)
  
Do While Not rs.EOF

  Set itmX = lsw.ListItems.Add(, , rs!consec)
      itmX.SubItems(1) = cbo.ItemData(cbo.ListIndex)
      itmX.SubItems(2) = rs!Cedula
      itmX.SubItems(3) = rs!Nombre
      itmX.SubItems(4) = rs!Cantidad
      itmX.SubItems(5) = Format(rs!Monto, "Standard")
      itmX.Tag = rs!Cod_Beneficio
            
  rs.MoveNext
Loop

rs.Close

End Sub

Private Sub cmdCerrar_Click()
    frDetalle.Visible = False

    cmdBuscar.Enabled = True
    cmdGenerar.Enabled = True
    
    lsw.Visible = True
    
End Sub

Private Sub cmdGenerar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim y As Integer, vConsec As Integer
Dim vBeneficio As String, vCedula As String

On Error GoTo vError


Me.MousePointer = vbHourglass

For i = 1 To lsw.ListItems.Count
    y = y + 1
Next i

If y > 0 Then

    prgBar.Max = y
    prgBar.Value = 1
    
    strSQL = ""
      
   For i = 1 To lsw.ListItems.Count
      If lsw.ListItems.Item(i).Checked Then
              vConsec = lsw.ListItems.Item(i).Text
              vBeneficio = lsw.ListItems.Item(i).SubItems(1)
              vCedula = lsw.ListItems.Item(i).SubItems(2)
              
              'Actualiza el estado en tabla afi_bene_otorga
              strSQL = strSQL & Space(10) & "Update afi_bene_otorga set estado = 'E',autoriza_user = '" & glogon.Usuario _
                     & "',autoriza_fecha = dbo.MyGetdate()" _
                     & " where cedula = '" & vCedula & "'" _
                     & " and cod_beneficio = '" & vBeneficio & "' and consec = '" & vConsec & "'"
                                      
              If Len(strSQL) > 20000 Then
                Call ConectionExecute(strSQL)
                strSQL = ""
              End If
           
        If prgBar < y Then
           prgBar.Value = prgBar.Value + 1
        End If
      End If
   Next i
  
End If

'Lote final
If Len(strSQL) > 0 Then
  Call ConectionExecute(strSQL)
End If

prgBar.Value = 0

Me.MousePointer = vbDefault

Call cmdBuscar_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

  Call cmdBuscar_Click
End Sub

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 7

imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Código", 1200, vbCenter
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 2800
    .Add , , "Cantidad", 1200, vbRightJustify
    .Add , , "Monto", 1600, vbRightJustify
End With

With lswDetalle.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Código", 1200, vbCenter
    .Add , , "Producto Id", 1800
    .Add , , "Producto", 2800
    .Add , , "Cantidad", 1200, vbRightJustify
    .Add , , "Costo/Ud", 1600, vbRightJustify
End With



strSQL = "select rtrim(cod_Beneficio) as 'Idx',  rtrim(descripcion) as 'ItmX'" _
       & " from afi_beneficios" _
       & "  where cod_beneficio in (select cod_beneficio from afi_bene_prodasg)"
           
Call sbCbo_Llena_New(cbo, strSQL, False, True)
              
Call cbo_Click
              
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim itmX As ListViewItem
Dim strSQL As String, rs As New ADODB.Recordset

lswDetalle.ListItems.Clear

If chkDetalle.Value = vbChecked Then
    cmdBuscar.Enabled = False
    cmdGenerar.Enabled = False
    
    frDetalle.Visible = True
    frDetalle.Caption = "Id. " & Item.Text & " [" & Item.SubItems(1) & "] " & Item.SubItems(3)
    
    lsw.Visible = False
    
    strSQL = "Select B.*, P.descripcion as 'ProductoDesc' " _
           & " from afi_bene_prodasg B inner join afi_bene_productos P on B.cod_producto = P.cod_Producto" _
           & " where consec = " & Item.Text & " and cod_beneficio = '" & Item.SubItems(1) & "'"
        
    Call OpenRecordSet(rs, strSQL)
    
    Do While Not rs.EOF
    
    Set itmX = lswDetalle.ListItems.Add(, , rs!consec)
        itmX.SubItems(1) = cbo.ItemData(cbo.ListIndex)
        itmX.SubItems(2) = rs!cod_producto
        itmX.SubItems(3) = rs!ProductoDesc
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = Format(rs!costo_unidad, "Standard")
    rs.MoveNext
    Loop
    
    rs.Close
End If

End Sub

