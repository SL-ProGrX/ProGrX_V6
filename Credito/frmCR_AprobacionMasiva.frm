VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCR_AprobacionMasiva 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aprobación Masiva"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   13710
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   13335
      _Version        =   1441792
      _ExtentX        =   23521
      _ExtentY        =   9975
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   7800
      Width           =   13335
      _Version        =   1441792
      _ExtentX        =   23521
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Resumen:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   360
         Width           =   1695
         _Version        =   1441792
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Marcar Todos?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtCasos 
         Height          =   315
         Left            =   5040
         TabIndex        =   13
         Top             =   360
         Width           =   1095
         _Version        =   1441792
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   315
         Left            =   7680
         TabIndex        =   14
         Top             =   360
         Width           =   2415
         _Version        =   1441792
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   375
         Left            =   10440
         TabIndex        =   16
         Top             =   360
         Width           =   1455
         _Version        =   1441792
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Formalizar"
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_AprobacionMasiva.frx":0000
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   12
         Top             =   360
         Width           =   975
         _Version        =   1441792
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   975
         _Version        =   1441792
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Casos"
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
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ProgressBar Prgbar 
      Height          =   150
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   13335
      _Version        =   1441792
      _ExtentX        =   23521
      _ExtentY        =   265
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   11040
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_AprobacionMasiva.frx":0727
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   661
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripción 
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8493
      _ExtentY        =   661
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label lblTitulo 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   480
      Width           =   8655
      _Version        =   1441792
      _ExtentX        =   15266
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aprobación Masiva de Créditos"
      ForeColor       =   16777215
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Línea de Crédito"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Rango"
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
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmCR_AprobacionMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHeight As Long, mWidth As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vPaso As Boolean



Private Sub btnAplicar_Click()
On Error GoTo vError

Dim i As Long, pProcess As Integer

Me.MousePointer = vbHourglass

Prgbar.Visible = True

With lsw.ListItems
    Prgbar.Max = .Count + 1
    Prgbar.Value = 0
    
    pProcess = 0
    strSQL = ""
    For i = 1 To .Count
        If .Item(i).Checked Then
           strSQL = strSQL & Space(10) & "exec spCrd_AprobacionMasiva_Formaliza " & .Item(i).Text & ", '" & glogon.Usuario & "'"
           
           pProcess = pProcess + 1
           
        End If
    
        Prgbar.Value = Prgbar.Value + 1
    Next i
    
    If pProcess >= 3 Then
        pProcess = 0
        Call ConectionExecute(strSQL)
        
        strSQL = ""
    End If
    
End With


If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

Prgbar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Operaciones Procesadas Satisfactoriamente!", vbInformation

Call btnBuscar_Click

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spCrd_AprobacionMasiva_Consulta '" & txtCodigo.Text _
        & "', '" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00','" & Format(dtpCorte.Value, "yyyy-mm-dd") _
        & " 23:59:59','" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

txtCasos.Text = 0
txtTotal.Text = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_Solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = Format(rs!Fecha_Solicita, "yyyy-mm-dd")
     itmX.SubItems(5) = Format(rs!Monto, "Standard")
     itmX.SubItems(6) = rs!Plazo
     itmX.SubItems(7) = Format(rs!Tasa, "Standard")
     itmX.SubItems(8) = Format(rs!Cuota, "Standard")
     itmX.SubItems(9) = rs!Garantia_Desc
     itmX.SubItems(10) = rs!Linea_Desc
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

vPaso = False

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkTodos_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass


Dim pCasos As Long, pTotal As Currency
Dim i As Long

pCasos = 0
pTotal = 0

With lsw.ListItems
    For i = 1 To lsw.ListItems.Count
        .Item(i).Checked = chkTodos.Value
        
        If .Item(i).Checked Then
               pTotal = pTotal + CCur(.Item(i).SubItems(5))
               pCasos = pCasos + 1
        End If
        
    Next i
End With

txtCasos.Text = Format(pCasos, "###,###0")
txtTotal.Text = Format(pTotal, "Standard")


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

mHeight = 13830
mWidth = 8280


With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 1800
    .Add , , "Código", 1100, vbCenter
    .Add , , "Identificación", 1800, vbCenter
    .Add , , "Nombre", 3800
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Monto", 2100, vbRightJustify
    .Add , , "Plazo", 1000, vbRightJustify
    .Add , , "Tasa", 1000, vbRightJustify
    .Add , , "Cuota", 2100, vbRightJustify
    .Add , , "Garantía", 2100
    .Add , , "Línea Desc", 3100
End With

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub Form_Resize()
On Error Resume Next

Dim pHeight As Long, pWidth As Long

If Me.Height < mHeight Then
   pHeight = Me.Height 'mHeight
Else
   pHeight = Me.Height
End If


If Me.Width < mWidth Then
   pWidth = mWidth
Else
   pWidth = Me.Width
End If

Prgbar.Width = pWidth - (Prgbar.Left + 200)
lsw.Width = Prgbar.Width


lsw.Height = pHeight - (lsw.top + 650 + gbResumen.Height)

gbResumen.top = lsw.top + lsw.Height + 100
gbResumen.Width = lsw.Width


End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim pCasos As Long, pTotal As Currency

pCasos = txtCasos.Text
pTotal = txtTotal.Text

If Item.Checked Then
    pCasos = pCasos + 1
    pTotal = pTotal + CCur(Item.SubItems(5))
Else
    pCasos = pCasos - 1
    pTotal = pTotal - CCur(Item.SubItems(5))

End If

txtCasos.Text = Format(pCasos, "###,###0")
txtTotal.Text = Format(pTotal, "Standard")


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Columna = "Codigo"
        gBusquedas.Orden = "Codigo"
        gBusquedas.Filtro = " and (Retencion = 'N' and Poliza = 'N')"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
            txtCodigo.Text = gBusquedas.Resultado
            txtDescripción.Text = gBusquedas.Resultado2
        End If
End If
End Sub
