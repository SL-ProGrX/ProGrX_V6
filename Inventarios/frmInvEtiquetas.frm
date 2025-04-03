VERSION 5.00
Begin VB.Form frmInvEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Etiquetas de Códigos de Barras"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   8430
   Begin VB.ComboBox cboRedondeo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtRedondeo 
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
      Left            =   1200
      TabIndex        =   13
      Text            =   "1"
      Top             =   2520
      Width           =   855
   End
   Begin VB.Data DaoX 
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdGeneraSato 
      Caption         =   "&Generación Sato"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
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
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   38
      TabIndex        =   8
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtFactura 
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtProvDesc 
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
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   38
      TabIndex        =   3
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox txtProvCod 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton opt 
      Caption         =   "Por Entradas (Compra de Mercadería)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.OptionButton opt 
      Caption         =   "Por Articulo (Producto / Servicio)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdGeneraSIFC 
      Caption         =   "&Generación Interna"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Redondeo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblProducto 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblFactura 
      Caption         =   "# Factura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblProv 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   8280
      X2              =   120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   8280
      X2              =   120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   8280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   8280
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmInvEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdGeneraSato_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curRedondeo As Currency


On Error Resume Next
'Borrar Archivo y Copia Original
Kill "C:\SIFBarra.dbf"
FileCopy (App.Path & "\SIFBarra.dbf"), "C:\SIFBarra.dbf"


Me.MousePointer = vbHourglass

On Error GoTo vError

If Mid(cboRedondeo, 1, 1) = "D" Then
 curRedondeo = txtRedondeo
Else
 curRedondeo = txtRedondeo * -1
End If

DaoX.DatabaseName = "C:\"
DaoX.RecordSource = "SIFBarra.dbf"
DaoX.Refresh
With DaoX.Recordset
 Do While Not .EOF
   .Delete
   .MoveNext
 Loop

    Select Case True
      Case opt.Item(0) 'Solo un producto
         strSQL = "select 1 as Cantidad,modelo,cod_barras,descripcion,cod_producto" _
                & ",round((precio_regular * ((impuesto_ventas / 100)+1))," & curRedondeo & ") as Precio" _
                & " from pv_productos where cod_producto = '" & txtCodigo & "'"
      Case opt.Item(1) 'Una Compra
         strSQL = "select E.Cantidad,P.modelo,P.cod_barras,P.descripcion,P.cod_producto" _
                & ",round((precio_regular * ((impuesto_ventas / 100)+1))," & curRedondeo & ") as Precio" _
                & " from pv_productos P inner join Cpr_Compras_detalle E On P.cod_producto = E.cod_producto" _
                & " where E.cod_proveedor = " & txtProvCod & " and cod_factura = '" & txtFactura & "'"
    End Select
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     For i = 1 To CLng(rs!cantidad)
        .AddNew
        !Codigo = fxBarrasProceso(rs!cod_producto)
        !NOMBRE1 = Mid(rs!Descripcion, 1, 30)
        !NOMBRE2 = rs!Modelo
        !Precio = Format(rs!Precio, "Standard")
        .Update
     Next i
     rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault
MsgBox "Archivo de Barras Actualizado...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxBarrasProceso(vProducto As String) As String
Dim strSQL As String, i As Integer, xBarra As String
Dim rs As New ADODB.Recordset

strSQL = "select cod_barras,cod_ProdClas from pv_productos where cod_producto = '" & vProducto & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  xBarra = rs!cod_barras & ""
  xBarra = Trim(xBarra)
  If Len(xBarra) < 12 Then
    xBarra = "2000" & Mid(Format(Trim(rs!COD_PRODCLAS), "000"), 1, 3) & Mid(Format(Trim(vProducto), "00000"), 1, 5)
    strSQL = "update pv_productos set cod_barras = '" & xBarra & "' where cod_producto = '" & vProducto & "'"
    Call ConectionExecute(strSQL)
  End If

   fxBarrasProceso = xBarra

Else
    fxBarrasProceso = ""
End If
rs.Close

End Function


Private Sub cmdGeneraSIFC_Click()
MsgBox "En Desarrollo..."
End Sub



Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()
vModulo = 32

cboRedondeo.AddItem "Decimales"
cboRedondeo.AddItem "Enteros"
cboRedondeo.Text = "Decimales"

Call Formularios(Me)
Call RefrescaTags(Me)

Call opt_Click(0)
End Sub

Private Sub opt_Click(Index As Integer)
 txtCodigo.Enabled = False
 txtNombre.Enabled = False
 lblProducto.ForeColor = vbBlack
 txtProvCod.Enabled = False
 txtProvDesc.Enabled = False
 txtFactura.Enabled = False
 lblProv.ForeColor = vbBlack
 lblFactura.ForeColor = vbBlack

Select Case True
  Case opt.Item(0) 'Producto
    txtCodigo.Enabled = True
    txtNombre.Enabled = True
    lblProducto.ForeColor = vbBlue
  Case opt.Item(1) 'Compra
    txtProvCod.Enabled = True
    txtProvDesc.Enabled = True
    txtFactura.Enabled = True
    lblProv.ForeColor = vbBlue
    lblFactura.ForeColor = vbBlue
End Select

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
  frmBusquedaArticulos.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodigo_LostFocus()
txtNombre = fxSIFCCodigos("D", txtCodigo, "Productos")
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "Select cod_producto,descripcion from pv_productos"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = ""
  gBusquedas.Convertir = "N"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "Select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = ""
  gBusquedas.Convertir = "N"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "Select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = ""
  gBusquedas.Convertir = "N"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If
End Sub

