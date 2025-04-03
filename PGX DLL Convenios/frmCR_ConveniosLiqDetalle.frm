VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_ConveniosLiqDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCR_ConveniosLiqDetalle.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   14025
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6252
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   12972
      _Version        =   1310722
      _ExtentX        =   22881
      _ExtentY        =   11028
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
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton BtnArchivo 
      Height          =   312
      Left            =   7080
      TabIndex        =   8
      Top             =   7320
      Width           =   852
      _Version        =   1310722
      _ExtentX        =   1503
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "&Archivo"
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
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   7680
      Top             =   120
   End
   Begin XtremeSuiteControls.PushButton btnExcel 
      Height          =   312
      Left            =   7920
      TabIndex        =   10
      Top             =   7320
      Width           =   852
      _Version        =   1310722
      _ExtentX        =   1503
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "&Excel"
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
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   1920
      TabIndex        =   11
      Top             =   7320
      Width           =   4932
      _Version        =   1310722
      _ExtentX        =   8705
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboInforme 
      Height          =   312
      Left            =   1920
      TabIndex        =   12
      Top             =   6960
      Width           =   1692
      _Version        =   1310722
      _ExtentX        =   2990
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal_1 
      Height          =   312
      Left            =   5040
      TabIndex        =   13
      Top             =   6960
      Width           =   1812
      _Version        =   1310722
      _ExtentX        =   3196
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
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal_2 
      Height          =   312
      Left            =   8400
      TabIndex        =   14
      Top             =   6960
      Width           =   1812
      _Version        =   1310722
      _ExtentX        =   3196
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
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal_3 
      Height          =   312
      Left            =   11400
      TabIndex        =   15
      Top             =   6960
      Width           =   1812
      _Version        =   1310722
      _ExtentX        =   3196
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
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1200
      TabIndex        =   16
      Top             =   120
      Width           =   972
      _Version        =   1310722
      _ExtentX        =   1714
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOrden 
      Height          =   312
      Left            =   9120
      TabIndex        =   17
      Top             =   120
      Width           =   1212
      _Version        =   1310722
      _ExtentX        =   2138
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   312
      Left            =   11520
      TabIndex        =   18
      Top             =   120
      Width           =   1692
      _Version        =   1310722
      _ExtentX        =   2984
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   2160
      TabIndex        =   19
      Top             =   120
      Width           =   5772
      _Version        =   1310722
      _ExtentX        =   10181
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblInstitucion 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   7320
      Width           =   1572
   End
   Begin VB.Label lblTotal_3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10560
      TabIndex        =   6
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblInforme 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Informe"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblTotal_2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   4
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lblTotal_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   10560
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Orden"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   8160
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmCR_ConveniosLiqDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mConvenio As String, mOrden As Long, mTipo As String
Dim vPaso As Boolean, mLoad_Inicial As Boolean

Private Sub BtnArchivo_Click()
    Select Case mTipo
    
       Case "Cargos"
       
       Case "Recaudacion", "ComisionRecaudacion"
            Call sbRecaudacion_Archivo
        
       Case "Devoluciones", "ComisionDevolucion"
            Call sbDevolucion_Archivo
        
       Case "NuevosCreditos", "ComisionCreditos"
            Call sbCreditos_Archivo
    
       Case "Reservas"
          
    End Select
End Sub

Private Sub btnExcel_Click()
Call Excel_Exportar_Lsw(lsw)
End Sub

Private Sub cboInforme_Click()

If vPaso Then Exit Sub
Call sbConsultas

End Sub


Private Sub cboInstitucion_Click()
If vPaso Then Exit Sub
Call sbConsultas
End Sub

Private Sub Form_Activate()
 vModulo = 16
End Sub

Private Sub Form_Load()
 
 vModulo = 16
  
 mConvenio = GLOBALES.gTag
 mOrden = GLOBALES.gTag2
 mTipo = GLOBALES.gTag3
 
 mLoad_Inicial = True
 
 vPaso = True
    cboInforme.AddItem "Resumen"
    cboInforme.AddItem "Detalle"
    cboInforme.AddItem "Institución"
    cboInforme.Text = "Resumen"
 vPaso = False
 
 
 txtTotal_2.Visible = False
 txtTotal_3.Visible = False
 lblTotal_2.Visible = False
 lblTotal_3.Visible = False
 
 Call sbConsultaConvenio
 
End Sub

Private Sub sbConsultaConvenio()
Dim strSQL As String, rs As New ADODB.Recordset
  
On Error GoTo vError

   strSQL = " select O.COD_CONVENIO,C.DESCRIPCION,O.COD_ORDEN,O.ESTADO" _
          & " from CRD_CONVENIOS_ORDENES O  inner join CRD_CONVENIOS C on O.COD_CONVENIO = C.COD_CONVENIO" _
          & " where O.COD_CONVENIO = '" & mConvenio & "' AND O.COD_ORDEN = " & mOrden & " "
   Call OpenRecordSet(rs, strSQL)

   If Not rs.EOF Then
      txtCodigo.Text = rs!COD_CONVENIO
      txtDescripcion.Text = rs!Descripcion
      txtOrden.Text = rs!cod_orden
      
      Select Case rs!estado
        Case "A"
          txtEstado.Text = "Abierta"
        Case "C"
          txtEstado.Text = "Cerrada"
        Case "T"
          txtEstado.Text = "Tramitada"
      End Select
   End If

   rs.Close
   
   
   
  'Consulta Inicial por Institucion
  vPaso = True
    Select Case mTipo
    
       Case "Cargos"
            strSQL = "exec spConvenios_Orden_Creditos_Institucion_Rsm '" & mConvenio & "'," & mOrden
       
       Case "Recaudacion", "ComisionRecaudacion"
            strSQL = "exec spConvenios_Orden_Recaudacion_Institucion_Rsm '" & mConvenio & "'," & mOrden
       
       Case "Devoluciones", "ComisionDevolucion"
            strSQL = "exec spConvenios_Orden_Devolucion_Institucion_Rsm '" & mConvenio & "'," & mOrden
       
       Case "NuevosCreditos", "ComisionCreditos"
            strSQL = "exec spConvenios_Orden_Creditos_Institucion_Rsm '" & mConvenio & "'," & mOrden
    
       Case "Reservas"
            strSQL = "exec spConvenios_Orden_Creditos_Institucion_Rsm '" & mConvenio & "'," & mOrden
          
    End Select
  
  
    Call OpenRecordSet(rs, strSQL)
    cboInstitucion.Clear
    cboInstitucion.AddItem "TODAS"
    cboInstitucion.Text = "TODAS"
    
    Do While Not rs.EOF
        cboInstitucion.AddItem Trim(rs!Institucion)
        cboInstitucion.ItemData(cboInstitucion.NewIndex) = rs!Codigo
     rs.MoveNext
    Loop
    rs.Close
  
  vPaso = False
  
Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsultas()
Dim strSQL As String, rs As New ADODB.Recordset
  
On Error GoTo vError

  Select Case mTipo
  
     Case "Cargos"
        Call sbCargos
     
     Case "Recaudacion", "ComisionRecaudacion"
        Call sbRecaudacion
     
     Case "Devoluciones", "ComisionDevolucion"
        Call sbDevolucion
     
     Case "NuevosCreditos", "ComisionCreditos"
        Call sbCreditosNuevos
        
     Case "Reservas"
        Call sbRecaudacion
        
  End Select
    
Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Sub

'Calcula el monto por cargos
Private Sub sbCargos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency

Me.Caption = "Cargos de Formalización Asociados al Convenio"

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

txtTotal_1.Text = "0.00"
lblTotal_1.Caption = "T.Recaudado"

curTotal = 0

If Mid(cboInforme.Text, 1, 1) = "R" Then
    lsw.ColumnHeaders.Add , , "Cargo", 700
    lsw.ColumnHeaders.Add , , "Descripción", 3000
    lsw.ColumnHeaders.Add , , "Monto", 1200, 1
 
    strSQL = "exec spConvenios_Orden_Cargos_Formaliza '" & mConvenio & "'," & mOrden & ",'R'"
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!cod_cargo)
         itmX.SubItems(1) = rs!Descripcion
         itmX.SubItems(2) = Format(rs!cargo, "Standard")
      
      curTotal = curTotal + rs!cargo
      
      rs.MoveNext
    Loop
    rs.Close
Else
  'Detalle
    
    lsw.ColumnHeaders.Add , , "Cargo", 700
    lsw.ColumnHeaders.Add , , "Descripción", 3000
    lsw.ColumnHeaders.Add , , "Recaudado", 1200, vbRightJustify
    lsw.ColumnHeaders.Add , , "Operación", 1200
    lsw.ColumnHeaders.Add , , "Línea", 3000
    lsw.ColumnHeaders.Add , , "Garantía", 1200
    lsw.ColumnHeaders.Add , , "Identificación", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Nombre", 3200
    lsw.ColumnHeaders.Add , , "Mnt. Crd", 1300
    lsw.ColumnHeaders.Add , , "Mnt. Girado.", 1300, vbRightJustify
    lsw.ColumnHeaders.Add , , "Formalizado", 1300, vbRightJustify
    lsw.ColumnHeaders.Add , , "Revisado", 1300
    lsw.ColumnHeaders.Add , , "Línea.Cod.", 1200
    
 
    strSQL = "exec spConvenios_Orden_Cargos_Formaliza '" & mConvenio & "'," & mOrden & ",'D'"
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!cod_cargo)
         itmX.SubItems(1) = rs!Descripcion
         itmX.SubItems(2) = Format(rs!cargo, "Standard")
         itmX.SubItems(3) = rs!id_Solicitud
         itmX.SubItems(4) = rs!LineaDesc
         itmX.SubItems(5) = rs!Garantia
         itmX.SubItems(6) = rs!Cedula
         itmX.SubItems(7) = rs!Nombre
         itmX.SubItems(8) = rs!montoapr
         itmX.SubItems(9) = rs!monto_girado
         itmX.SubItems(10) = Format(rs!fechaforp, "dd/mm/yyyy")
         itmX.SubItems(11) = Format(rs!Fecha_Revision, "dd/mm/yyyy")
         itmX.SubItems(12) = rs!Codigo
      
      
      curTotal = curTotal + rs!cargo
      
      rs.MoveNext
    Loop
    rs.Close
  
End If



txtTotal_1.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

'Calcula el total de los recaudaciones de creditos
Private Sub sbRecaudacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency, curComision As Currency

Dim vInstitucion As String, vInstMonto As Currency, vInstComision As Currency

Me.Caption = "Recaudaciones"

On Error GoTo vError

Me.MousePointer = vbHourglass


lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

txtTotal_2.Visible = True
lblTotal_2.Visible = True

txtTotal_3.Visible = True
lblTotal_3.Visible = True


txtTotal_1.Text = "0.00"
lblTotal_1.Caption = "Recaudado"

lblTotal_2.Caption = "Comisión"
txtTotal_2.Text = "0.00"

lblTotal_3.Caption = "Neto"
txtTotal_3.Text = "0.00"


curTotal = 0
curComision = 0

Select Case Mid(cboInforme.Text, 1, 1)
  Case "R"
    lsw.ColumnHeaders.Add , , "Línea", 1500
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Recaudado", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Comisión", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Neto", 2500, vbRightJustify
    lsw.ColumnHeaders.Add , , "Institución", 3500
 
    vInstitucion = "Inicial"
    vInstMonto = 0
    vInstComision = 0
    
    
    If cboInstitucion.Text = "TODAS" Then
        strSQL = "exec spConvenios_Orden_Recaudacion '" & mConvenio & "'," & mOrden & ",'R'"
    Else
        strSQL = "exec spConvenios_Orden_Recaudacion_Institucion '" & mConvenio & "'," & mOrden & ",'R'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If

    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
        If vInstitucion = "Inicial" Then
           vInstitucion = Trim(rs!Institucion)
        End If
    
           
        If vInstitucion <> Trim(rs!Institucion) Then
        
           Set itmX = lsw.ListItems.Add(, , "")
               itmX.SubItems(1) = Space(20) & "Total:"
               itmX.SubItems(2) = Format(vInstMonto, "Standard")
               itmX.SubItems(3) = Format(vInstComision, "Standard")
               itmX.SubItems(4) = Format(vInstMonto - vInstComision, "Standard")
               itmX.SubItems(5) = Trim(vInstitucion)
               
               itmX.ListSubItems.Item(1).Bold = True
               itmX.ListSubItems.Item(2).Bold = True
               itmX.ListSubItems.Item(3).Bold = True
               itmX.ListSubItems.Item(4).Bold = True
               itmX.ListSubItems.Item(5).Bold = True
               
               itmX.ListSubItems.Item(5).ForeColor = vbBlue
           Set itmX = lsw.ListItems.Add(, , "")
           
           vInstitucion = Trim(rs!Institucion)
           vInstMonto = 0
           vInstComision = 0
        End If
        
      
        Set itmX = lsw.ListItems.Add(, , rs!Codigo)
           itmX.SubItems(1) = Trim(rs!LineaX)
           itmX.SubItems(2) = Format(rs!TOTAL_RECAUDADO, "Standard")
           itmX.SubItems(3) = Format(rs!Comision, "Standard")
           itmX.SubItems(4) = Format(rs!TOTAL_NETO, "Standard")
'           itmX.SubItems(5) = Trim(rs!Institucion)
      
      
        vInstMonto = vInstMonto + rs!TOTAL_RECAUDADO
        vInstComision = vInstComision + rs!Comision
      
        curTotal = curTotal + rs!TOTAL_RECAUDADO
        curComision = curComision + rs!Comision
      
       rs.MoveNext
       If rs.EOF Then
           Set itmX = lsw.ListItems.Add(, , "")
               itmX.SubItems(1) = Space(20) & "Total:"
               itmX.SubItems(2) = Format(vInstMonto, "Standard")
               itmX.SubItems(3) = Format(vInstComision, "Standard")
               itmX.SubItems(4) = Format(vInstMonto - vInstComision, "Standard")
               itmX.SubItems(5) = Trim(vInstitucion)
               itmX.ListSubItems.Item(1).Bold = True
               itmX.ListSubItems.Item(2).Bold = True
               itmX.ListSubItems.Item(3).Bold = True
               itmX.ListSubItems.Item(4).Bold = True
               itmX.ListSubItems.Item(5).Bold = True
               
               itmX.ListSubItems.Item(5).ForeColor = vbBlue
        End If
    Loop
    rs.Close
    
  Case "D" 'Detalle
    lsw.ColumnHeaders.Add , , "Operación", 1200
    lsw.ColumnHeaders.Add , , "Línea", 1200
    lsw.ColumnHeaders.Add , , "Recaudado", 1200, vbRightJustify
    lsw.ColumnHeaders.Add , , "Fecha", 1200, vbRightJustify
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Garantía", 1200
    lsw.ColumnHeaders.Add , , "Identificación", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Nombre", 3200
    lsw.ColumnHeaders.Add , , "Concepto", 3200
    lsw.ColumnHeaders.Add , , "Tipo Doc.", 2200
    lsw.ColumnHeaders.Add , , "Documento", 2200
    lsw.ColumnHeaders.Add , , "Institución", 3500
    
 
    If cboInstitucion.Text = "TODAS" Then
        strSQL = "exec spConvenios_Orden_Recaudacion '" & mConvenio & "'," & mOrden & ",'D'"
    Else
        strSQL = "exec spConvenios_Orden_Recaudacion_Institucion '" & mConvenio & "'," & mOrden & ",'D'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!id_Solicitud)
         itmX.SubItems(1) = rs!Codigo
         itmX.SubItems(2) = Format(rs!Monto, "Standard")
         itmX.SubItems(3) = Format(rs!fecha, "dd/mm/yyyy")
         itmX.SubItems(4) = Trim(rs!LineaX)
         itmX.SubItems(5) = Trim(rs!GarantiaDesc)
         itmX.SubItems(6) = Trim(rs!Cedula)
         itmX.SubItems(7) = Trim(rs!Nombre)
         itmX.SubItems(8) = rs!concepto
         itmX.SubItems(9) = rs!Tcon
         itmX.SubItems(10) = rs!Ncon
         itmX.SubItems(11) = rs!Institucion
        
        curTotal = curTotal + rs!Monto
        curComision = curComision + rs!Comision
      
      rs.MoveNext
    Loop
    rs.Close
  
  Case "I" 'Institucion
    lsw.ColumnHeaders.Add , , "Código", 1500
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Recaudado", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Comisión", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Neto", 2500, vbRightJustify
    
    
    strSQL = "exec spConvenios_Orden_Recaudacion_Institucion_Rsm '" & mConvenio & "'," & mOrden
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
        Set itmX = lsw.ListItems.Add(, , rs!Codigo)
           itmX.SubItems(1) = Trim(rs!Institucion)
           itmX.SubItems(2) = Format(rs!TOTAL_RECAUDADO, "Standard")
           itmX.SubItems(3) = Format(rs!Comision, "Standard")
           itmX.SubItems(4) = Format(rs!TOTAL_NETO, "Standard")
        
        curTotal = curTotal + rs!TOTAL_RECAUDADO
        curComision = curComision + rs!Comision
      
      rs.MoveNext
    Loop
    rs.Close
  
  
  
End Select

txtTotal_1.Text = Format(curTotal, "Standard")
txtTotal_2.Text = Format(curComision, "Standard")
txtTotal_3.Text = Format(curTotal - curComision, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbDevolucion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency, curComision As Currency

Dim vInstitucion As String, vInstMonto As Currency, vInstComision As Currency

Me.Caption = "Devoluciones"

On Error GoTo vError

Me.MousePointer = vbHourglass


lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

txtTotal_2.Visible = True
lblTotal_2.Visible = True

txtTotal_3.Visible = True
lblTotal_3.Visible = True


txtTotal_1.Text = "0.00"
lblTotal_1.Caption = "Recaudado"

lblTotal_2.Caption = "Comisión"
txtTotal_2.Text = "0.00"

lblTotal_3.Caption = "Neto"
txtTotal_3.Text = "0.00"


curTotal = 0
curComision = 0

Select Case Mid(cboInforme.Text, 1, 1)
  Case "R"
    lsw.ColumnHeaders.Add , , "Concepto", 1500
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Recaudado", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Comisión", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Neto", 2500, vbRightJustify
    lsw.ColumnHeaders.Add , , "Institución", 3500
 
    vInstitucion = "Inicial"
    vInstMonto = 0
    vInstComision = 0
    
    
    If cboInstitucion.Text = "TODAS" Then
        strSQL = "exec spConvenios_Orden_Devolucion '" & mConvenio & "'," & mOrden & ",'R'"
    Else
        strSQL = "exec spConvenios_Orden_Devolucion_Institucion '" & mConvenio & "'," & mOrden & ",'R'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If

    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
        If vInstitucion = "Inicial" Then
           vInstitucion = Trim(rs!Institucion)
        End If
    
           
        If vInstitucion <> Trim(rs!Institucion) Then
        
           Set itmX = lsw.ListItems.Add(, , "")
               itmX.SubItems(1) = Space(20) & "Total:"
               itmX.SubItems(2) = Format(vInstMonto, "Standard")
               itmX.SubItems(3) = Format(vInstComision, "Standard")
               itmX.SubItems(4) = Format(vInstMonto - vInstComision, "Standard")
               itmX.SubItems(5) = Trim(vInstitucion)
               
               itmX.ListSubItems.Item(1).Bold = True
               itmX.ListSubItems.Item(2).Bold = True
               itmX.ListSubItems.Item(3).Bold = True
               itmX.ListSubItems.Item(4).Bold = True
               itmX.ListSubItems.Item(5).Bold = True
               
               itmX.ListSubItems.Item(5).ForeColor = vbBlue
           Set itmX = lsw.ListItems.Add(, , "")
           
           vInstitucion = Trim(rs!Institucion)
           vInstMonto = 0
           vInstComision = 0
        End If
        
      
        Set itmX = lsw.ListItems.Add(, , rs!Retencion_Codigo)
           itmX.SubItems(1) = Trim(rs!RetencionConcepto)
           itmX.SubItems(2) = Format(rs!TOTAL_RECAUDADO, "Standard")
           itmX.SubItems(3) = Format(rs!Comision, "Standard")
           itmX.SubItems(4) = Format(rs!TOTAL_NETO, "Standard")
'           itmX.SubItems(5) = Trim(rs!Institucion)
      
      
        vInstMonto = vInstMonto + rs!TOTAL_RECAUDADO
        vInstComision = vInstComision + rs!Comision
      
        curTotal = curTotal + rs!TOTAL_RECAUDADO
        curComision = curComision + rs!Comision
      
       rs.MoveNext
       If rs.EOF Then
           Set itmX = lsw.ListItems.Add(, , "")
               itmX.SubItems(1) = Space(20) & "Total:"
               itmX.SubItems(2) = Format(vInstMonto, "Standard")
               itmX.SubItems(3) = Format(vInstComision, "Standard")
               itmX.SubItems(4) = Format(vInstMonto - vInstComision, "Standard")
               itmX.SubItems(5) = Trim(vInstitucion)
               itmX.ListSubItems.Item(1).Bold = True
               itmX.ListSubItems.Item(2).Bold = True
               itmX.ListSubItems.Item(3).Bold = True
               itmX.ListSubItems.Item(4).Bold = True
               itmX.ListSubItems.Item(5).Bold = True
               
               itmX.ListSubItems.Item(5).ForeColor = vbBlue
        End If
    Loop
    rs.Close
    
  Case "D" 'Detalle
    lsw.ColumnHeaders.Add , , "Num.Liq.", 1200
    lsw.ColumnHeaders.Add , , "Concepto", 3200
    lsw.ColumnHeaders.Add , , "Recaudado", 1200, vbRightJustify
    lsw.ColumnHeaders.Add , , "Fecha", 1200, vbRightJustify
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Identificación", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Nombre", 3200
    lsw.ColumnHeaders.Add , , "Operadora", 3200
    lsw.ColumnHeaders.Add , , "Plan", 2200
    lsw.ColumnHeaders.Add , , "Contrato", 2200
    lsw.ColumnHeaders.Add , , "Institución", 3500
    
 
    If cboInstitucion.Text = "TODAS" Then
        strSQL = "exec spConvenios_Orden_Devolucion '" & mConvenio & "'," & mOrden & ",'D'"
    Else
        strSQL = "exec spConvenios_Orden_Devolucion_Institucion '" & mConvenio & "'," & mOrden & ",'D'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!Consec)
         itmX.SubItems(1) = rs!Retencion_Codigo
         itmX.SubItems(2) = Format(rs!Monto, "Standard")
         itmX.SubItems(3) = Format(rs!fecha, "dd/mm/yyyy")
         itmX.SubItems(4) = Trim(rs!RetencionConcepto)
         itmX.SubItems(5) = Trim(rs!Cedula)
         itmX.SubItems(6) = Trim(rs!Nombre)
         itmX.SubItems(7) = rs!cod_Operadora
         itmX.SubItems(8) = rs!PlanDesc
         itmX.SubItems(9) = rs!cod_Contrato
         itmX.SubItems(10) = rs!Institucion
        
        curTotal = curTotal + rs!Monto
        curComision = curComision + rs!Comision
      
      rs.MoveNext
    Loop
    rs.Close
  
  Case "I" 'Institucion
    lsw.ColumnHeaders.Add , , "Código", 1500
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Recaudado", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Comisión", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Neto", 2500, vbRightJustify
    
    
    strSQL = "exec spConvenios_Orden_Devolucion_Institucion_Rsm '" & mConvenio & "'," & mOrden
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
        Set itmX = lsw.ListItems.Add(, , rs!Codigo)
           itmX.SubItems(1) = Trim(rs!Institucion)
           itmX.SubItems(2) = Format(rs!TOTAL_RECAUDADO, "Standard")
           itmX.SubItems(3) = Format(rs!Comision, "Standard")
           itmX.SubItems(4) = Format(rs!TOTAL_NETO, "Standard")
        
        curTotal = curTotal + rs!TOTAL_RECAUDADO
        curComision = curComision + rs!Comision
      
      rs.MoveNext
    Loop
    rs.Close
  
  
  
End Select

txtTotal_1.Text = Format(curTotal, "Standard")
txtTotal_2.Text = Format(curComision, "Standard")
txtTotal_3.Text = Format(curTotal - curComision, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbRecaudacion_Archivo()
Dim strSQL As String, rs As New ADODB.Recordset

Dim vTempo As String, vArchivo As String, vFile As String, vCadena As String, vRuta As String
Dim fnFile

On Error GoTo vError

Me.MousePointer = vbHourglass

On Error Resume Next


'Crea Directorios
fnFile = FreeFile

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Convenios"

vRuta = SIFGlobal.DirectorioDeResultados & "\Convenios"
vArchivo = "Recaudacion_" & mConvenio & "_" & Format(mOrden, "0000") & " [" & Mid(cboInstitucion.Text, 1, 20) & "].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError

If cboInstitucion.Text = "TODAS" Then
    strSQL = "exec spConvenios_Orden_Recaudacion '" & mConvenio & "'," & mOrden & ",'D'"
Else
    strSQL = "exec spConvenios_Orden_Recaudacion_Institucion '" & mConvenio & "'," & mOrden & ",'D'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If
Call OpenRecordSet(rs, strSQL)


Open vTempo For Output As #fnFile  ' Create file name.

vCadena = "Cédula,Nombre,Monto,Fecha,Comprobante,Institucion,Referencia"
Print #fnFile, vCadena

Do While Not rs.EOF

  vCadena = rs!Cedula & "," & rs!Nombre _
        & "," & rs!Monto _
        & "," & Format(rs!fecha, "dd/mm/yyyy") _
        & "," & rs!Tcon & "-" & rs!Ncon _
        & "," & rs!Institucion _
        & "," & rs!Codigo & "." & rs!id_Solicitud
 Print #fnFile, vCadena

 rs.MoveNext
Loop
rs.Close

Close #fnFile

Me.MousePointer = vbDefault
MsgBox "Archivo CSV Generado para Abrir en Excel:" & vTempo, vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbDevolucion_Archivo()
Dim strSQL As String, rs As New ADODB.Recordset

Dim vTempo As String, vArchivo As String, vFile As String, vCadena As String, vRuta As String
Dim fnFile

On Error GoTo vError

Me.MousePointer = vbHourglass

On Error Resume Next


'Crea Directorios
fnFile = FreeFile

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Convenios"

vRuta = SIFGlobal.DirectorioDeResultados & "\Convenios"
vArchivo = "Devolucion_" & mConvenio & "_" & Format(mOrden, "0000") & " [" & Mid(cboInstitucion.Text, 1, 20) & "].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError

If cboInstitucion.Text = "TODAS" Then
    strSQL = "exec spConvenios_Orden_Devolucion '" & mConvenio & "'," & mOrden & ",'D'"
Else
    strSQL = "exec spConvenios_Orden_Devolucion_Institucion '" & mConvenio & "'," & mOrden & ",'D'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If
Call OpenRecordSet(rs, strSQL)


Open vTempo For Output As #fnFile  ' Create file name.

vCadena = "Cédula,Nombre,Monto,Fecha,Comprobante,Institucion,Referencia"
Print #fnFile, vCadena

Do While Not rs.EOF

  vCadena = rs!Cedula & "," & rs!Nombre _
        & "," & rs!Monto _
        & "," & Format(rs!fecha, "dd/mm/yyyy") _
        & ",Liq." & rs!Consec _
        & "," & rs!Institucion _
        & "," & rs!Retencion_Codigo & "." & rs!cod_Plan & "." & rs!cod_Contrato
 Print #fnFile, vCadena

 rs.MoveNext
Loop
rs.Close

Close #fnFile

Me.MousePointer = vbDefault
MsgBox "Archivo CSV Generado para Abrir en Excel:" & vTempo, vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCreditos_Archivo()
Dim strSQL As String, rs As New ADODB.Recordset

Dim vTempo As String, vArchivo As String, vFile As String, vCadena As String, vRuta As String
Dim fnFile

On Error GoTo vError

Me.MousePointer = vbHourglass

On Error Resume Next


'Crea Directorios
fnFile = FreeFile

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Convenios"

vRuta = SIFGlobal.DirectorioDeResultados & "\Convenios"
vArchivo = "CreditosNuevos_" & mConvenio & "_" & Format(mOrden, "0000") & " [" & Mid(cboInstitucion.Text, 1, 20) & "].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError


If cboInstitucion.Text = "TODAS" Then
    strSQL = "exec spConvenios_Orden_Creditos '" & mConvenio & "'," & mOrden & ",'D'"
Else
    strSQL = "exec spConvenios_Orden_Creditos_Institucion '" & mConvenio & "'," & mOrden & ",'D'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If
 
Call OpenRecordSet(rs, strSQL)


Open vTempo For Output As #fnFile  ' Create file name.

vCadena = "Cédula,Nombre,Monto,Comision,Neto,Fecha,Documento,Institucion,Operación,Linea,Destino"
Print #fnFile, vCadena

Do While Not rs.EOF

  vCadena = rs!Cedula & "," & rs!Nombre _
        & "," & rs!Monto & "," & rs!Comision & "," & rs!Neto _
        & "," & Format(rs!fechaforp, "dd/mm/yyyy") _
        & "," & Trim(rs!Documento & "") _
        & "," & rs!InstitucionDesc _
        & "," & rs!id_Solicitud & "," & rs!Codigo & "," & rs!Destino
 Print #fnFile, vCadena

 rs.MoveNext
Loop
rs.Close

Close #fnFile

Me.MousePointer = vbDefault
MsgBox "Archivo CSV Generado para Abrir en Excel:" & vTempo, vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



'Calcula montos aprobados de nuevos créditos
Private Sub sbCreditosNuevos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency, curComision As Currency

Dim vInstitucion As String, vInstMonto As Currency, vInstComision As Currency

Me.Caption = "Créditos Formalizados"

On Error GoTo vError

Me.MousePointer = vbHourglass


lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

txtTotal_2.Visible = True
lblTotal_2.Visible = True

txtTotal_3.Visible = True
lblTotal_3.Visible = True


txtTotal_1.Text = "0.00"
lblTotal_1.Caption = "T.Girado"

lblTotal_2.Caption = "Comisión"
txtTotal_2.Text = "0.00"

lblTotal_3.Caption = "Neto"
txtTotal_3.Text = "0.00"


curTotal = 0
curComision = 0

   
   
Select Case Mid(cboInforme.Text, 1, 1)
  Case "R" 'Resumen
    lsw.ColumnHeaders.Add , , "Línea", 1500
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Monto", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Comisión", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Neto", 2500, vbRightJustify
    lsw.ColumnHeaders.Add , , "Destino", 3500
    lsw.ColumnHeaders.Add , , "Institución", 3500
 
 
    vInstitucion = "Inicial"
    vInstMonto = 0
    vInstComision = 0
    
    
    If cboInstitucion.Text = "TODAS" Then
        strSQL = "exec spConvenios_Orden_Creditos '" & mConvenio & "'," & mOrden & ",'R'"
    Else
        strSQL = "exec spConvenios_Orden_Creditos_Institucion '" & mConvenio & "'," & mOrden & ",'R'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If
    
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
        If vInstitucion = "Inicial" Then
           vInstitucion = Trim(rs!Institucion)
        End If
        
        If vInstitucion <> Trim(rs!Institucion) Then
        
           Set itmX = lsw.ListItems.Add(, , "")
               itmX.SubItems(1) = Space(20) & "Total:"
               itmX.SubItems(2) = Format(vInstMonto, "Standard")
               itmX.SubItems(3) = Format(vInstComision, "Standard")
               itmX.SubItems(4) = Format(vInstMonto - vInstComision, "Standard")
               itmX.SubItems(6) = Trim(vInstitucion)
               
               itmX.ListSubItems.Item(1).Bold = True
               itmX.ListSubItems.Item(2).Bold = True
               itmX.ListSubItems.Item(3).Bold = True
               itmX.ListSubItems.Item(4).Bold = True
               itmX.ListSubItems.Item(6).Bold = True
               
               itmX.ListSubItems.Item(6).ForeColor = vbBlue
           
           Set itmX = lsw.ListItems.Add(, , "")
           
           vInstitucion = Trim(rs!Institucion)
           vInstMonto = 0
           vInstComision = 0
        End If
        
        
        Set itmX = lsw.ListItems.Add(, , rs!Codigo)
           itmX.SubItems(1) = Trim(rs!LineaDesc)
           itmX.SubItems(2) = Format(rs!Monto, "Standard")
           itmX.SubItems(3) = Format(rs!Comision, "Standard")
           itmX.SubItems(4) = Format(rs!Neto, "Standard")
           itmX.SubItems(5) = Trim(rs!Destino)
           itmX.SubItems(6) = Trim(rs!Institucion)
        
               
               
        vInstMonto = vInstMonto + rs!Monto
        vInstComision = vInstComision + rs!Comision
        
        curTotal = curTotal + rs!Monto
        curComision = curComision + rs!Comision
      
      rs.MoveNext
      
      If rs.EOF Then
           Set itmX = lsw.ListItems.Add(, , "")
               itmX.SubItems(1) = Space(20) & "Total:"
               itmX.SubItems(2) = Format(vInstMonto, "Standard")
               itmX.SubItems(3) = Format(vInstComision, "Standard")
               itmX.SubItems(4) = Format(vInstMonto - vInstComision, "Standard")
               itmX.SubItems(6) = Trim(vInstitucion)
               
               itmX.ListSubItems.Item(1).Bold = True
               itmX.ListSubItems.Item(2).Bold = True
               itmX.ListSubItems.Item(3).Bold = True
               itmX.ListSubItems.Item(4).Bold = True
               itmX.ListSubItems.Item(6).Bold = True
               
               itmX.ListSubItems.Item(6).ForeColor = vbBlue
           
      End If
      
    Loop
    rs.Close

 Case "D"
 
  'Detalle
    lsw.ColumnHeaders.Add , , "Operación", 1200
    lsw.ColumnHeaders.Add , , "Línea", 1000, vbCenter
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Destino", 3000
    lsw.ColumnHeaders.Add , , "Monto", 2000, vbRightJustify
    lsw.ColumnHeaders.Add , , "Comisión", 2000, vbRightJustify
    lsw.ColumnHeaders.Add , , "Neto", 2000, vbRightJustify
    lsw.ColumnHeaders.Add , , "Fec.Rev.", 2000
    lsw.ColumnHeaders.Add , , "Fec.Form.", 2000
    lsw.ColumnHeaders.Add , , "No.Documento", 2000
    lsw.ColumnHeaders.Add , , "Garantía", 2000
    lsw.ColumnHeaders.Add , , "Identificación", 1800, vbCenter
    lsw.ColumnHeaders.Add , , "Nombre", 3200
    lsw.ColumnHeaders.Add , , "Institución", 3500
    
    If cboInstitucion.Text = "TODAS" Then
        strSQL = "exec spConvenios_Orden_Creditos '" & mConvenio & "'," & mOrden & ",'D'"
    Else
        strSQL = "exec spConvenios_Orden_Creditos_Institucion '" & mConvenio & "'," & mOrden & ",'D'," & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If
 
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!id_Solicitud)
         itmX.SubItems(1) = rs!Codigo
         itmX.SubItems(2) = Trim(rs!LineaDesc)
         itmX.SubItems(3) = Trim(rs!Destino)
         itmX.SubItems(4) = Format(rs!Monto, "Standard")
         itmX.SubItems(5) = Format(rs!Comision, "Standard")
         itmX.SubItems(6) = Format(rs!Neto, "Standard")
         itmX.SubItems(7) = Format(rs!Fecha_Revision, "dd/mm/yyyy")
         itmX.SubItems(8) = Format(rs!fechaforp, "dd/mm/yyyy")
         itmX.SubItems(9) = Trim(rs!Documento & "")
         itmX.SubItems(10) = Trim(rs!Garantia)
         itmX.SubItems(11) = Trim(rs!Cedula)
         itmX.SubItems(12) = Trim(rs!Nombre)
         itmX.SubItems(13) = Trim(rs!InstitucionDesc)
        
        curTotal = curTotal + rs!Monto
        curComision = curComision + rs!Comision
      
      rs.MoveNext
    Loop
    rs.Close
  
 Case "I"
    lsw.ColumnHeaders.Add , , "Código", 1500
    lsw.ColumnHeaders.Add , , "Descripción", 3500
    lsw.ColumnHeaders.Add , , "Monto", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Comisión", 2400, vbRightJustify
    lsw.ColumnHeaders.Add , , "Neto", 2500, vbRightJustify
 
 
    vInstitucion = "Inicial"
    vInstMonto = 0
    vInstComision = 0
    
    
    strSQL = "exec spConvenios_Orden_Creditos_Institucion_Rsm '" & mConvenio & "'," & mOrden
    
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
       
        Set itmX = lsw.ListItems.Add(, , rs!cod_Institucion)
           itmX.SubItems(1) = Trim(rs!Institucion)
           itmX.SubItems(2) = Format(rs!Monto, "Standard")
           itmX.SubItems(3) = Format(rs!Comision, "Standard")
           itmX.SubItems(4) = Format(rs!Neto, "Standard")
        
        curTotal = curTotal + rs!Monto
        curComision = curComision + rs!Comision
      
      rs.MoveNext
    Loop
    rs.Close
  
End Select

txtTotal_1.Text = Format(curTotal, "Standard")
txtTotal_2.Text = Format(curComision, "Standard")
txtTotal_3.Text = Format(curTotal - curComision, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'  Resume
End Sub

Private Sub Form_Resize()
On Error Resume Next

lsw.Height = Me.Height - (lsw.Top + 1600)
lsw.Width = Me.Width - (lsw.Left + 500)
lblInforme.Top = lsw.Top + lsw.Height + 100
cboInforme.Top = lblInforme.Top
lblTotal_1.Top = lblInforme.Top
txtTotal_1.Top = lblInforme.Top
lblTotal_2.Top = lblInforme.Top
txtTotal_2.Top = lblInforme.Top
lblTotal_3.Top = lblInforme.Top
txtTotal_3.Top = lblInforme.Top

cboInstitucion.Top = cboInforme.Top + cboInforme.Height + 100
lblInstitucion.Top = cboInstitucion.Top
BtnArchivo.Top = cboInstitucion.Top
btnExcel.Top = BtnArchivo.Top

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub Timer1_Timer()
Timer1.Interval = 0
Timer1.Enabled = False

 Call sbConsultas
End Sub
