VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSYS_Mercadito_Archivo_Load 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Interface: El Mercadito"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ComboBox cboProveedor 
      Height          =   312
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   432
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12086
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   0
      Left            =   8280
      TabIndex        =   3
      Top             =   1560
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Archivo_Load.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   1
      Left            =   8760
      TabIndex        =   4
      Top             =   1560
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Archivo_Load.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   2
      Left            =   9240
      TabIndex        =   5
      Top             =   1560
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Archivo_Load.frx":0E19
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   12015
      _Version        =   524288
      _ExtentX        =   21193
      _ExtentY        =   9551
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      SpreadDesigner  =   "frmSYS_Mercadito_Archivo_Load.frx":1532
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnImportar 
      Height          =   372
      Left            =   9960
      TabIndex        =   8
      Top             =   1560
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Importar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Archivo_Load.frx":1E2B
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   372
      Left            =   11280
      TabIndex        =   9
      Top             =   1560
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Archivo_Load.frx":251B
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Archivo"
      ForeColor       =   8388608
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Proveedor"
      ForeColor       =   8388608
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   612
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   4572
      _Version        =   1441793
      _ExtentX        =   8064
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Importación de Productos "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1050
      Left            =   0
      Picture         =   "frmSYS_Mercadito_Archivo_Load.frx":2C1B
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmSYS_Mercadito_Archivo_Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim strSQL As String, rs As New ADODB.Recordset
Dim pEmpresaId As Long


Private Sub sbLimpia()
On Error GoTo vError

    vGrid.MaxRows = 0
    txtArchivo.Text = ""

vError:
End Sub


Private Sub sbArchivo_Find()


With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        txtArchivo.Text = .FileName
    
End With

End Sub




Private Sub sbArchivo_Load()

Dim i As Long, vCampos As Boolean

Dim pNombre As String, pDescripcion As String, pModelo As String, pMarca As String, pCodigo As String
Dim pCategoria As Long, pCantidad As Long, pPrecioRack As Currency, pPrecio As Currency, pIVA As Currency
Dim pRetiroMaximo As Long, pRetiroDias As Long

On Error GoTo vError


vGrid.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboProveedor.ListCount <= 0 Then
    MsgBox "Indique un Proveedor válido que se encuentra asociado a su cuenta...", vbCritical
    Exit Sub
End If


Me.MousePointer = vbHourglass



    Set rs = Excel_Load(txtArchivo.Text, "Productos")
        
    'Validaciónn del Archivo
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "NOMBRE" Then
           vCampos = True
        End If
        If vCampos Then Exit For
    Next i
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "DESCRIPCION" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "IDCATEGORIA" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "MARCA" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "MODELO" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "CODIGO" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "CANTIDAD" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "PRECIO_RACK" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "PRECIO" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "IMPUESTO" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    'FIN: Validación del Archivo
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "RETIROMAXIMO" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    vCampos = False
    For i = 0 To rs.Fields.Count
        If UCase(LCase(rs.Fields(i).Name)) = "RETIRODIAS" Then
           vCampos = True
        End If
         If vCampos Then Exit For
    Next i
    
    If Not vCampos Then
       MsgBox "No coincide la estructura del archivo a cargar...", vbExclamation
       Exit Sub
    End If
    
    
'Sube, Revisa y Carga
With vGrid
    Do While Not rs.EOF
     If Not IsNull(rs!Nombre) Then
            pNombre = RTrim(rs!Nombre & "")
            pDescripcion = RTrim(rs!Descripcion & "")
            pMarca = RTrim(rs!Marca & "")
            pModelo = RTrim(rs!Modelo & "")
            pCodigo = RTrim(rs!Codigo & "")
            pCantidad = IIf(IsNull(rs!Cantidad), 1, rs!Cantidad)
            pPrecioRack = IIf(IsNull(rs!Precio_Rack), 1, rs!Precio_Rack)
            pPrecio = IIf(IsNull(rs!Precio), 1, rs!Precio)
            pIVA = IIf(IsNull(rs!Impuesto), 1, rs!Impuesto)
            pCategoria = IIf(IsNull(rs!idCategoria), 1, rs!idCategoria)
      
      
            pRetiroMaximo = IIf(IsNull(rs!RetiroMaximo), 1, rs!RetiroMaximo)
            pRetiroDias = IIf(IsNull(rs!RetiroDias), 1, rs!RetiroDias)
      
      
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .Text = pNombre
            .Col = 2
            .Text = pDescripcion
            .Col = 3
            .Text = CStr(pCategoria)
            .Col = 4
            .Text = pMarca
            .Col = 5
            .Text = pModelo
            .Col = 6
            .Text = pCodigo
            
            .Col = 7
            .Text = CStr(pCantidad)
            
            .Col = 8
            .Text = Format(pPrecioRack, "Standard")
      
            .Col = 9
            .Text = Format(pPrecio, "Standard")
            
            .Col = 10
            .Text = Format(pIVA, "Standard")
      
      
            .Col = 11
            .Text = CStr(pRetiroMaximo)
      
            .Col = 12
            .Text = CStr(pRetiroDias)
      
      rs.MoveNext
      
      End If 'Nulos
    Loop
    rs.Close

End With 'vGrid


Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub


Private Sub sbImportar()
Dim pNombre As String, pDescripcion As String, pModelo As String, pMarca As String, pCodigo As String
Dim pCategoria As Long, pCantidad As Long, pPrecioRack As Currency, pPrecio As Currency, pIVA As Currency
Dim pProveedor As Long, pRetiroMaximo As Long, pRetiroDias As Long

Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

pProveedor = cboProveedor.ItemData(cboProveedor.ListIndex)

With vGrid

strSQL = ""

For i = 1 To .MaxRows
  .Row = i
  .Col = 1
  pNombre = .Text
  .Col = 2
  pDescripcion = .Text
  .Col = 3
  pCategoria = .Text
  .Col = 4
  pMarca = .Text
  .Col = 5
  pModelo = .Text
  .Col = 6
  pCodigo = .Text
  .Col = 7
  pCantidad = .Text
  .Col = 8
  pPrecioRack = .Text
  .Col = 9
  pPrecio = .Text
  .Col = 10
  pIVA = .Text
  
  .Col = 11
  pRetiroMaximo = .Text
  
  .Col = 12
  pRetiroDias = .Text
   
strSQL = strSQL & Space(10) & "exec spProductos_Import " & pProveedor & ",'" & pNombre & "','" & pDescripcion _
       & "'," & pCategoria & ",'" & pMarca & "','" & pModelo & "','" & pCodigo _
       & "'," & pCantidad & "," & pPrecioRack & "," & pPrecio & "," & pIVA & ", " & pRetiroMaximo & ", " & pRetiroDias

  If Len(strSQL) > 20000 Then
    db.Execute strSQL
    strSQL = ""
  End If
 
Next i
End With
 
'Lote Final
If Len(strSQL) > 0 Then
  db.Execute strSQL
  strSQL = ""
End If

 
 
Me.MousePointer = vbDefault
MsgBox "Importación realizada satisfactoriamente... Registros Procesados :" & vGrid.MaxRows, vbInformation

Call sbLimpia

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub





Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String


        
Select Case Index
  
  Case 0 'buscar
  
        txtArchivo.Text = ""
       Call sbArchivo_Find
  
  Case 1 'Cargar
       Call sbArchivo_Load

    
  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: Productos" & vbCrLf _
              & " 3. Columnas.: Nombre, Descripcion, idCategoria, Marca, Modelo, Codigo, Cantidad, RetiroMaximo, RetiroDias, Precio_Rack, Precio, Impuesto"
     
     MsgBox vMensaje, vbInformation
         
End Select

End Sub

Private Sub btnCancelar_Click()
Call sbLimpia
End Sub

Private Sub btnImportar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen Productos para Importar...[verifique!]", vbExclamation
       Exit Sub
    End If
   
    Call sbImportar
End Sub

Private Sub Form_Load()

On Error GoTo vError


pEmpresaId = 2

strSQL = "select PORTAL_ID from sif_Empresa"
Call OpenRecordSet(rs, strSQL)
    pEmpresaId = rs!Portal_Id
rs.Close


'Establece Conexion
strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=progrx.centralus.cloudapp.azure.com" _
       & ";Database=ElMercadito;APP=PGX_APL_Access;tcp:progrx.centralus.cloudapp.azure.com" _
       & "," & SIFGlobal.PuertosDisponibles & ";"
       
db.ConnectionString = strSQL
db.Open , "31M3rcadit0", "#S0n+oFl*v3M4t3w1/*"


strSQL = "select idEmpresa from Empresa where ClienteAPL = " & pEmpresaId
rs.Open strSQL, db, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   pEmpresaId = rs!idEmpresa
Else
    pEmpresaId = -1
End If
rs.Close

strSQL = "select P.IdProveedor as 'IdX', P.Descripcion as 'ItmX'" _
    & "   from Proveedor P" _
    & "    inner join ProveedoresEmpresas Pe on P.IdProveedor = Pe.IdProveedor" _
    & "    inner join Empresa E on Pe.IdEmpresa = Pe.IdEmpresa" _
    & "  Where E.IdEmpresa = " & pEmpresaId
    
rs.Open strSQL, db, adOpenStatic
Do While Not rs.EOF
 cboProveedor.AddItem rs!itmX & ""
 cboProveedor.ItemData(cboProveedor.ListCount - 1) = CStr(rs!IdX)
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboProveedor.Text = rs!itmX & ""
End If
rs.Close

Call sbLimpia

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - (vGrid.Left + 150)
vGrid.Height = Me.Height - (vGrid.Top + 250)

End Sub
