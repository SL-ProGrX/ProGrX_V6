VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmInvMargenUtilidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Margenes de Utilidad "
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmInvMargenUtilidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7215
   Begin XtremeSuiteControls.GroupBox gbMargenes 
      Height          =   5172
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   6732
      _Version        =   1245185
      _ExtentX        =   11874
      _ExtentY        =   9123
      _StockProps     =   79
      Caption         =   "Cambio de Margenes y Precios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin VB.TextBox txtUtilidadGeneral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   11
         Text            =   "0"
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txtUtilidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2412
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   5292
         _ExtentX        =   9340
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1685
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5953
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Utilidad"
            Object.Width           =   1544
         EndProperty
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   252
         Left            =   1680
         TabIndex        =   17
         Top             =   3360
         Width           =   2412
         _Version        =   1245185
         _ExtentX        =   4254
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplicar a todos los precios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkPrecioRegular 
         Height          =   252
         Left            =   1680
         TabIndex        =   18
         Top             =   3600
         Width           =   2412
         _Version        =   1245185
         _ExtentX        =   4254
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Actualiza Precio Regular"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkAgregar 
         Height          =   852
         Left            =   120
         TabIndex        =   19
         Top             =   4320
         Width           =   3972
         _Version        =   1245185
         _ExtentX        =   7006
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Si no existe el Tipo de Precio en el Producto Crearlo y Aplicarle el Margen de Utilidad Indicada en la Lista de Precios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Utilidad Precio Regular"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   1320
         TabIndex        =   16
         Top             =   3960
         Width           =   1812
      End
      Begin VB.Label lblPrecioDesc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   2280
         TabIndex        =   15
         Top             =   480
         Width           =   3372
      End
      Begin VB.Label lblPrecioCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Precios / Indique la Utilidad para Cada Uno"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Marque Aquellos que desea Aplicar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   732
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   1680
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.RadioButton rbAccion 
      Height          =   372
      Index           =   0
      Left            =   3720
      TabIndex        =   6
      Top             =   1080
      Width           =   3132
      _Version        =   1245185
      _ExtentX        =   5524
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cambio de Margen de Utilidad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Value           =   -1  'True
      Alignment       =   1
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   168
      Left            =   360
      TabIndex        =   1
      Top             =   7500
      Visible         =   0   'False
      Width           =   4872
      _ExtentX        =   8599
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   612
      Left            =   5400
      TabIndex        =   2
      Top             =   7200
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmInvMargenUtilidad.frx":030A
   End
   Begin XtremeSuiteControls.ComboBox cboLinea 
      Height          =   276
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   4452
      _Version        =   1245185
      _ExtentX        =   7858
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboLineaSub 
      Height          =   276
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   4452
      _Version        =   1245185
      _ExtentX        =   7858
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.RadioButton rbAccion 
      Height          =   372
      Index           =   1
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   3132
      _Version        =   1245185
      _ExtentX        =   5524
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Actualiza Precio Segun Margen Actual"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Línea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
End
Attribute VB_Name = "frmInvMargenUtilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean




Private Sub cboLinea_Click()
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select COD_LINEA_SUB as 'IdX',  DESCRIPCION as 'ItmX'" _
    & " From PV_PROD_CLASIFICA_SUB where COD_PRODCLAS = " & cboLinea.ItemData(cboLinea.ListIndex)
Call sbCbo_Llena_New(cboLineaSub, strSQL, False, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkTodos_Click()
Dim i As Integer

On Error GoTo vError

If chkTodos = vbChecked Then
  
  For i = 1 To lsw.ListItems.Count
      lsw.ListItems.Item(i).SubItems(2) = Format(txtUtilidadGeneral, "Standard")
  Next i
  
Else
  
  For i = 1 To lsw.ListItems.Count
      lsw.ListItems.Item(i).SubItems(2) = Format(0, "Standard")
  Next i
  
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub CmdAplicar_Click()

If cboLinea.ListCount = 0 Or cboLineaSub.ListCount = 0 Then Exit Sub

Select Case True
  Case rbAccion.Item(0).Value  'Margenes y Precios
    Call sbMargenPrecio
    
  Case rbAccion.Item(1).Value 'Solo Precios
    Call sbPrecio

End Select


End Sub

Private Sub sbMargenPrecio()
Dim strSQL As String
Dim i As Integer, vPrecio As String, vUtilidad As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

prgBar.Visible = True

  
If chkPrecioRegular.Value = xtpChecked Then
  strSQL = "update pv_productos set precio_regular = costo_regular + (costo_regular * " _
         & CCur(txtUtilidadGeneral) & " / 100), porc_utilidad = " & CCur(txtUtilidadGeneral) _
         & " where estado = 'A' and cod_prodclas = " & cboLinea.ItemData(cboLinea.ListIndex) _
         & " and COD_LINEA_SUB = '" & cboLineaSub.ItemData(cboLineaSub.ListIndex) & "'"
  Call ConectionExecute(strSQL)
End If



For i = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(i).Checked Then
    vUtilidad = CCur(lsw.ListItems.Item(i).SubItems(2)) / 100
    vPrecio = lsw.ListItems.Item(i).Text

    strSQL = "update X set porc_utilidad = " & (vUtilidad * 100) _
           & ",X.Monto = P.costo_Regular + (P.costo_Regular * " & vUtilidad & ")" _
           & " from pv_productos P inner join pv_producto_precios X" _
           & " on P.cod_producto = X.cod_producto" _
           & " where P.estado = 'A' and P.cod_prodclas = " & cboLinea.ItemData(cboLinea.ListIndex) _
           & " and P.COD_LINEA_SUB = '" & cboLineaSub.ItemData(cboLineaSub.ListIndex) & "'" _
           & " and X.cod_precio = '" & vPrecio & "'"
 End If
Next i

prgBar.Visible = False

Me.MousePointer = vbDefault

MsgBox "Margenes de Utilidad Actualizados...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbPrecio()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vPrecio As String, vUtilidad As Currency
Dim strPrecios As String

On Error GoTo vError

Me.MousePointer = vbHourglass

prgBar.Visible = True

strPrecios = ""
For i = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(i).Checked Then
   If Len(strPrecios) = 0 Then
       strPrecios = "'" & Trim(lsw.ListItems.Item(i).Text) & "'"
   Else
       strPrecios = strPrecios & ",'" & Trim(lsw.ListItems.Item(i).Text) & "'"
   End If
 End If
Next i

If Len(strPrecios) = 0 Then strPrecios = "''"

If chkPrecioRegular.Value = xtpChecked Then
    strSQL = "update pv_productos set precio_regular = costo_regular + (costo_regular * porc_utilidad / 100)" _
           & " where estado = 'A' and cod_prodclas = " & cboLinea.ItemData(cboLinea.ListIndex) _
           & " and COD_LINEA_SUB = '" & cboLineaSub.ItemData(cboLineaSub.ListIndex) & "'"
    Call ConectionExecute(strSQL)
End If

strSQL = "update X set X.Monto = P.costo_Regular + (P.costo_Regular * X.porc_utilidad /100)" _
       & " from pv_productos P inner join pv_producto_precios X" _
       & " on P.cod_producto = X.cod_producto" _
       & " where P.estado = 'A' and P.cod_prodclas = " & cboLinea.ItemData(cboLinea.ListIndex) _
       & " and P.COD_LINEA_SUB = '" & cboLineaSub.ItemData(cboLineaSub.ListIndex) & "'" _
       & " and X.cod_precio in(" & strPrecios & ")"
Call ConectionExecute(strSQL)

prgBar.Visible = False

Me.MousePointer = vbDefault

MsgBox "Precios Actualizados!", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

vModulo = 32

Me.MousePointer = vbHourglass

'Carga Lineas
vPaso = True
    cboLinea.Clear
    strSQL = "select cod_prodclas as 'IdX',descripcion as 'ItmX' from PV_PROD_CLASIFICA"
    Call sbCbo_Llena_New(cboLinea, strSQL, False, True)
vPaso = False

Call cboLinea_Click

strSQL = "select * from pv_tipos_precios"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_precio)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = 0
 rs.MoveNext
Loop
rs.Close

Call Formularios(Me)
Call RefrescaTags(Me)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub lsw_Click()
If lsw.ListItems.Count > 0 Then
  lblPrecioCod.Caption = lsw.SelectedItem
  lblPrecioDesc.Caption = lsw.SelectedItem.SubItems(1)
  txtUtilidad = lsw.SelectedItem.SubItems(2)
  txtUtilidad.SetFocus
End If
End Sub

Private Sub rbAccion_Click(Index As Integer)
Select Case Index
  Case 0 'Margenes y Precios
    txtUtilidad.Locked = False
    txtUtilidadGeneral.Locked = False
    
  Case 1 'Solo Precios
    txtUtilidad.Locked = True
    txtUtilidadGeneral.Locked = True

End Select
End Sub

Private Sub txtUtilidad_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

On Error GoTo vError

'Guardar el Precio
If KeyCode = vbKeyReturn And lblPrecioCod.Caption <> "" Then
  
  For i = 1 To lsw.ListItems.Count
    If Trim(lsw.ListItems.Item(i).Text) = Trim(lblPrecioCod.Caption) Then
      lsw.ListItems.Item(i).SubItems(2) = Format(txtUtilidad, "Standard")
      Exit For
    End If
  Next i
  
  lblPrecioCod.Caption = ""
  lblPrecioDesc.Caption = ""
  txtUtilidad = ""
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
