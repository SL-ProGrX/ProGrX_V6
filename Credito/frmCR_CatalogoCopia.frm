VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCR_CatalogoCopia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copia de Configuración : Línea "
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lswLineas 
      Height          =   4212
      Left            =   5760
      TabIndex        =   4
      Top             =   840
      Width           =   5412
      _Version        =   1441792
      _ExtentX        =   9546
      _ExtentY        =   7429
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4212
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   3492
      _Version        =   1441792
      _ExtentX        =   6159
      _ExtentY        =   7429
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   10932
      _Version        =   1441792
      _ExtentX        =   19283
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnCopiar 
         Height          =   372
         Left            =   8520
         TabIndex        =   8
         Top             =   360
         Width           =   1452
         _Version        =   1441792
         _ExtentX        =   2561
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Copiar"
         BackColor       =   -2147483633
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
         Picture         =   "frmCR_CatalogoCopia.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevaCodigo 
         Height          =   312
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevaDescripcion 
         Height          =   312
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   5412
         _Version        =   1441792
         _ExtentX        =   9546
         _ExtentY        =   550
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
         Appearance      =   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Línea"
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
         Index           =   3
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   1092
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   360
      Top             =   2280
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8880
      TabIndex        =   1
      Top             =   240
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtLinea 
      Height          =   312
      Left            =   2280
      TabIndex        =   9
      Top             =   240
      Width           =   852
      _Version        =   1441792
      _ExtentX        =   1503
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   5652
      _Version        =   1441792
      _ExtentX        =   9970
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A ...:"
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
      Index           =   2
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copiar ...:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea Base"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmCR_CatalogoCopia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vScroll As Boolean

Private Sub sbInicializa()

lsw.ListItems.Clear
lswLineas.ListItems.Clear

lsw.ListItems.Add , "0x01", "Configuración General"
lsw.ListItems.Add , "0x02", "Configuración Contable"
lsw.ListItems.Add , "0x03", "Destinos Asignados"
lsw.ListItems.Add , "0x04", "Cargos Asignados"
lsw.ListItems.Add , "0x05", "Recursos Presupuestarios"
lsw.ListItems.Add , "0x06", "Requisitos Asignados"
lsw.ListItems.Add , "0x07", "Cartera de Cobro Asociada"
lsw.ListItems.Add , "0x08", "Rangos (Montos,Plazos,Tasas)"
lsw.ListItems.Add , "0x09", "Niveles de Resolución"
lsw.ListItems.Add , "0x10", "Lista de Refundibles"
lsw.ListItems.Add , "0x11", "Adjuntos para Auto Gestión"


strSQL = "select codigo,descripcion from catalogo where activo = 1"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswLineas.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!Descripcion

 rs.MoveNext
Loop
rs.Close

End Sub

Private Sub btnCopiar_Click()
Dim x As Integer, y As Long, strSQL As String
Dim pGeneral As Byte, pCuentas As Byte, pRango As Byte, pDestinos As Byte, pRefundibles As Byte
Dim pCargos As Byte, pRecursos As Byte, pRequisitos As Byte, pCobro As Byte, pResolucion As Byte
Dim pAdjuntos As Byte

Me.MousePointer = vbHourglass

On Error GoTo vError


With lswLineas.ListItems
 For y = 1 To .Count
   If .Item(y).Checked And UCase(Trim(.Item(y).Text)) <> UCase(Trim(txtLinea)) Then
      
     pGeneral = 0
     pCuentas = 0
     pRango = 0
     pDestinos = 0
     pCargos = 0
     pRecursos = 0
     pRequisitos = 0
     pCobro = 0
     pResolucion = 0
     pRefundibles = 0
     pAdjuntos = 0
     
      For x = 1 To lsw.ListItems.Count
        If lsw.ListItems.Item(x).Checked Then
           Select Case lsw.ListItems.Item(x).Key
            Case "0x01" 'Configuración General
                pGeneral = 1
            Case "0x02" 'Configuración Contable
                pCuentas = 1
            Case "0x03" 'Destinos Asignados
                pDestinos = 1
            Case "0x04" 'Cargos Asignados
                pCargos = 1
            Case "0x05" 'Recursos Presuestarios
                pRecursos = 1
            Case "0x06" 'Requisitos Asignados
                pRequisitos = 1
            Case "0x07" 'Cartera de Cobro Asociada
                pCobro = 1
            Case "0x08" 'Rangos (Montos,Plazos,Tasas)
                pRango = 1
            Case "0x09" 'Niveles de Resolucion
                pResolucion = 1
            Case "0x10" 'Lista de Refundibles
                pRefundibles = 1
            Case "0x11" 'Adjuntos de Auto Gestion
                pAdjuntos = 1
           End Select
        
        End If
      Next x

      strSQL = "exec spCrdLineaCreditoCopia '" & txtLinea.Text & "','" & .Item(y).Text & "','" & glogon.Usuario & "'," & pGeneral _
             & "," & pCuentas & "," & pRango & "," & pDestinos & "," & pCargos & "," & pRecursos & "," & pRequisitos & "," & pCobro _
             & ",''," & pResolucion & "," & pRefundibles & "," & pAdjuntos
      Call ConectionExecute(strSQL)
      
      Call Bitacora("Aplica", "Copia Línea.:" & txtLinea.Text & " a " & .Item(y).Text & " (Cg:" & pGeneral & " Cc:" & pCuentas & " Da:" & pDestinos _
                  & " Ca:" & pCargos & " Rp:" & pRecursos & " Ra:" & pRequisitos & " CCa:" & pCobro & " Rmpt:" & pRango _
                  & " NivRes:" & pResolucion & " Lref:" & pRefundibles & " Adj:" & pAdjuntos & ")")
      
   End If
 Next y
End With

'Inserta linea
If txtNuevaCodigo.Text <> "" Then
     pGeneral = 0
     pCuentas = 0
     pRango = 0
     pDestinos = 0
     pCargos = 0
     pRecursos = 0
     pRequisitos = 0
     pCobro = 0
     pResolucion = 0
     pRefundibles = 0
     pAdjuntos = 0
          
      For x = 1 To lsw.ListItems.Count
        If lsw.ListItems.Item(x).Checked Then
           Select Case lsw.ListItems.Item(x).Key
            Case "0x01" 'Configuración General
                pGeneral = 1
            Case "0x02" 'Configuración Contable
                pCuentas = 1
            Case "0x03" 'Destinos Asignados
                pDestinos = 1
            Case "0x04" 'Cargos Asignados
                pCargos = 1
            Case "0x05" 'Recursos Presuestarios
                pRecursos = 1
            Case "0x06" 'Requisitos Asignados
                pRequisitos = 1
            Case "0x07" 'Cartera de Cobro Asociada
                pCobro = 1
            Case "0x08" 'Rangos (Montos,Plazos,Tasas)
                pRango = 1
            Case "0x09" 'Niveles de Resolucion
                pResolucion = 1
            Case "0x10" 'Lista de Refundibles
                pRefundibles = 1
            Case "0x11" 'Adjuntos de Auto Gestion
                pAdjuntos = 1
           End Select
        
        End If
      Next x

      strSQL = "exec spCrdLineaCreditoCopia '" & txtLinea.Text & "','" & txtNuevaCodigo.Text & "','" & glogon.Usuario & "'," & pGeneral _
             & "," & pCuentas & "," & pRango & "," & pDestinos & "," & pCargos & "," & pRecursos & "," & pRequisitos & "," & pCobro _
             & ",'" & txtNuevaDescripcion.Text & "'," & pResolucion & "," & pRefundibles & "," & pAdjuntos
      Call ConectionExecute(strSQL)

      Call Bitacora("Aplica", "Copia Línea.:" & txtLinea.Text & " a " & txtNuevaCodigo.Text & " (Cg:" & pGeneral & " Cc:" & pCuentas & " Da:" & pDestinos _
                  & " Ca:" & pCargos & " Rp:" & pRecursos & " Ra:" & pRequisitos & " CCa:" & pCobro & " Rmpt:" & pRango _
                  & " NivRes:" & pResolucion & " Lref:" & pRefundibles & " Adj:" & pAdjuntos & ")")

End If

Me.MousePointer = vbDefault
MsgBox "Copia Realizada Satisfactoriamente...", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 codigo from catalogo"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where retencion = 'N' and poliza = 'N' and  codigo > '" & txtLinea.Text & "' order by codigo asc"
    Else
       strSQL = strSQL & " where retencion = 'N' and poliza = 'N' and codigo < '" & txtLinea.Text & "' order by codigo desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtLinea.Text = rs!Codigo
      txtLinea_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
 
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "", 3500
End With
 
With lswLineas.ColumnHeaders
    .Clear
    .Add , , "Código", 1000
    .Add , , "Descripción", 3500
End With
 
 
 
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lswLineas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswLineas.SortKey = ColumnHeader.Index - 1
  If lswLineas.SortOrder = 0 Then lswLineas.SortOrder = 1 Else lswLineas.SortOrder = 0
  lswLineas.Sorted = True
End Sub


Private Sub Timer1_Timer()

Timer1.Interval = 0
Call sbInicializa

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Filtro = " and retencion = 'N' and poliza = 'N'"
  gBusquedas.Orden = "codigo"
  frmBusquedas.Show vbModal, Me
  If gBusquedas.Resultado <> "" Then
    txtLinea.Text = gBusquedas.Resultado
    txtDescripcion.Text = gBusquedas.Resultado2
  End If
End If

End Sub

Private Sub txtLinea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "codigo"
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Filtro = " and retencion = 'N' and poliza = 'N'"
  gBusquedas.Orden = "codigo"
  frmBusquedas.Show vbModal, Me
  If gBusquedas.Resultado <> "" Then
    txtLinea.Text = gBusquedas.Resultado
    txtDescripcion.Text = gBusquedas.Resultado2
  End If
End If

End Sub

Private Sub txtLinea_LostFocus()
  txtDescripcion.Text = fxDescribeCodigo(txtLinea.Text)
End Sub
