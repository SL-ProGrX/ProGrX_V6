VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvExistenciaProducto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Existencia x Producto al Corte"
   ClientHeight    =   6888
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   8472
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6888
   ScaleWidth      =   8472
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lswExistencia 
      Height          =   5652
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   8172
      _Version        =   1245187
      _ExtentX        =   14414
      _ExtentY        =   9970
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   16
      Picture         =   "frmInvExistenciaProducto.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInvCorte 
      Height          =   312
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigoX 
      Height          =   312
      Left            =   1080
      TabIndex        =   3
      Top             =   240
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
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombreX 
      Height          =   312
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   5292
      _Version        =   1245187
      _ExtentX        =   9334
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
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Index           =   1
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cerrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   16
      Picture         =   "frmInvExistenciaProducto.frx":0700
      ImageAlignment  =   4
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmInvExistenciaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuscar_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curSum As Currency

Select Case Index
  Case 0 'buscar
        lswExistencia.ListItems.Clear
        curSum = 0
        
        strSQL = "select cod_bodega,descripcion from pv_bodegas "
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Set itmX = lswExistencia.ListItems.Add(, , rs!cod_bodega)
              itmX.SubItems(1) = rs!Descripcion
              itmX.SubItems(2) = Format(fxInvProcesoProd(txtCodigoX, rs!cod_bodega, dtpInvCorte.Value), "###,###,###,##0")
              curSum = curSum + CCur(itmX.SubItems(2))
          rs.MoveNext
        Loop
        rs.Close
          
        
        Set itmX = lswExistencia.ListItems.Add(, , "")
            itmX.SubItems(2) = "______________"
        
        Set itmX = lswExistencia.ListItems.Add(, , "Total")
            itmX.SubItems(2) = Format(curSum, "###,###,###,##0")
       
   
   
  Case 1 'cerrar
    Unload Me
    
End Select

End Sub

Private Sub Form_Load()
    
With lswExistencia.ColumnHeaders
        .Clear
        .Add , , "Bodega", 1200
        .Add , , "Descripción", 4000
        .Add , , "Existencia", 1600, vbCenter
End With
    
dtpInvCorte.Value = fxFechaServidor

txtCodigoX = ""
txtNombreX = ""

End Sub


Private Sub txtCodigoX_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombreX.SetFocus
If KeyCode = vbKeyF4 Then
  frmBusquedaArticulos.Show vbModal
  txtCodigoX = gBusquedas.Resultado
  txtNombreX = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCodigoX_LostFocus()
txtNombreX = fxSIFCCodigos("D", txtCodigoX, "Productos")
End Sub

Private Sub txtNombreX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInvCorte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_producto,descripcion from pv_productos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigoX = gBusquedas.Resultado
  txtNombreX = gBusquedas.Resultado2
End If
End Sub



