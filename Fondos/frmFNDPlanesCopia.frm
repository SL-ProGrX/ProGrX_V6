VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDPlanesCopia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copia de Configuración de Planes"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4212
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   3492
      _Version        =   1572864
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
   Begin XtremeSuiteControls.ListView lswPlanes 
      Height          =   4212
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   5412
      _Version        =   1572864
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
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   240
      Top             =   1560
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   10932
      _Version        =   1572864
      _ExtentX        =   19283
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnCopiar 
         Height          =   375
         Left            =   9000
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Copiar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmFNDPlanesCopia.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoCodigo 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtNuevoDescripcion 
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   5415
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1092
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9240
      TabIndex        =   7
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   556
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
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   240
      Width           =   5655
      _Version        =   1572864
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
      Caption         =   "Plan Base"
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
      TabIndex        =   12
      Top             =   240
      Width           =   1092
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
      TabIndex        =   11
      Top             =   840
      Width           =   1092
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
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmFNDPlanesCopia"
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
lswPlanes.ListItems.Clear

lsw.ListItems.Add , "0x01", "Configuración General"
lsw.ListItems.Add , "0x02", "Configuración Contable"
lsw.ListItems.Add , "0x03", "Política de Puntos Add"
lsw.ListItems.Add , "0x04", "Política de Multas"
lsw.ListItems.Add , "0x05", "Destinos de los Ahorros"
lsw.ListItems.Add , "0x06", "Estados de la Persona"
lsw.ListItems.Add , "0x07", "Plazos de Vencimientos"



strSQL = "select COD_OPERADORA, COD_PLAN, DESCRIPCION, ESTADO" _
      & "  From FND_PLANES where Estado = 'A' Order by DESCRIPCION"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswPlanes.ListItems.Add(, , rs!Cod_Plan)
     itmX.SubItems(1) = rs!Descripcion
 rs.MoveNext
Loop
rs.Close

End Sub

Private Sub btnCopiar_Click()
Dim x As Integer, y As Long, strSQL As String

Dim pGeneral As Byte, pCuentas As Byte, pPlazos As Byte, pDestinos As Byte, pRetiros As Byte
Dim pEstadoPersona As Byte, pPtsAdd As Byte

Me.MousePointer = vbHourglass

On Error GoTo vError


With lswPlanes.ListItems
 For y = 1 To .Count
   If .Item(y).Checked And UCase(Trim(.Item(y).Text)) <> UCase(Trim(txtCodigo)) Then
      
     pGeneral = 0
     pCuentas = 0
     pPlazos = 0
     pDestinos = 0
     pRetiros = 0
     pEstadoPersona = 0
     pPtsAdd = 0
     
    
      For x = 1 To lsw.ListItems.Count
        If lsw.ListItems.Item(x).Checked Then
           Select Case lsw.ListItems.Item(x).Key
            Case "0x01" 'Configuración General
                pGeneral = 1
            Case "0x02" 'Configuración Contable
                pCuentas = 1
            Case "0x03" 'Política de Puntos Add
                pPtsAdd = 1
            Case "0x04" 'Política de Multas
                pRetiros = 1
            Case "0x05" 'Destinos de los Ahorros
                pDestinos = 1
            Case "0x06" 'Estados de la Persona
                pEstadoPersona = 1
            Case "0x07" 'Plazos de Vencimientos
                pPlazos = 1
           End Select
        
        End If
      Next x
      

      strSQL = "exec spFndPlanesCopia 0, '" & txtCodigo.Text & "','" & .Item(y).Text & "','" & glogon.Usuario _
             & "', " & pRetiros & ", " & pPtsAdd & ", " & pGeneral & ", " & pCuentas & ", " & pDestinos _
             & ",  " & pEstadoPersona & ", " & pPlazos & ",''"
      Call ConectionExecute(strSQL)
      
      Call Bitacora("Aplica", "Copia Plan.:" & txtCodigo.Text & " a " & .Item(y).Text & " (Cg:" & pGeneral & " Cc:" & pCuentas & " Da:" & pDestinos _
                  & " PTs:" & pPtsAdd & " Est:" & pEstadoPersona & " Plz:" & pPlazos & " Multas:" & pRetiros & ")")
      
   End If
 Next y
End With

'Inserta linea
If txtNuevoCodigo.Text <> "" Then
     pGeneral = 0
     pCuentas = 0
     pPlazos = 0
     pDestinos = 0
     pRetiros = 0
     pEstadoPersona = 0
     pPtsAdd = 0
          
     txtNuevoCodigo.Text = fxSysCleanTxtInject(txtNuevoCodigo.Text)
     txtNuevoDescripcion.Text = fxSysCleanTxtInject(txtNuevoDescripcion.Text)
          
      For x = 1 To lsw.ListItems.Count
        If lsw.ListItems.Item(x).Checked Then
           Select Case lsw.ListItems.Item(x).Key
            Case "0x01" 'Configuración General
                pGeneral = 1
            Case "0x02" 'Configuración Contable
                pCuentas = 1
            Case "0x03" 'Política de Puntos Add
                pPtsAdd = 1
            Case "0x04" 'Política de Multas
                pRetiros = 1
            Case "0x05" 'Destinos de los Ahorros
                pDestinos = 1
            Case "0x06" 'Estados de la Persona
                pEstadoPersona = 1
            Case "0x07" 'Plazos de Vencimientos
                pPlazos = 1
           End Select
        
        End If
      Next x

      strSQL = "exec spFndPlanesCopia 0, '" & txtCodigo.Text & "','" & txtNuevoCodigo.Text & "','" & glogon.Usuario _
             & "', " & pRetiros & ", " & pPtsAdd & ", " & pGeneral & ", " & pCuentas & ", " & pDestinos _
             & ",  " & pEstadoPersona & ", " & pPlazos & ",'" & txtNuevoDescripcion.Text & "'"
      Call ConectionExecute(strSQL)
      
      Call Bitacora("Aplica", "Copia Plan.:" & txtCodigo.Text & " a " & txtNuevoCodigo.Text & " (Cg:" & pGeneral & " Cc:" & pCuentas & " Da:" & pDestinos _
                  & " PTs:" & pPtsAdd & " Est:" & pEstadoPersona & " Plz:" & pPlazos & " Multas:" & pRetiros & ")")
      


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
    strSQL = "select Top 1 cod_Plan from fnd_Planes"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Estado = 'A' and cod_Plan > '" & txtCodigo.Text & "' order by cod_Plan asc"
    Else
       strSQL = strSQL & " where Estado = 'A' and cod_Plan < '" & txtCodigo.Text & "' order by cod_Plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Cod_Plan
      txtCodigo_LostFocus
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
 
vModulo = 18

txtCodigo.Text = GLOBALES.gTag2
txtDescripcion.Text = GLOBALES.gTag3


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "", 3500
End With
 
With lswPlanes.ColumnHeaders
    .Clear
    .Add , , "Código", 1000
    .Add , , "Descripción", 3500
End With
 
 
 
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lswPlanes_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswPlanes.SortKey = ColumnHeader.Index - 1
  If lswPlanes.SortOrder = 0 Then lswPlanes.SortOrder = 1 Else lswPlanes.SortOrder = 0
  lswPlanes.Sorted = True
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
  gBusquedas.Consulta = "select cod_Plan,descripcion from fnd_Planes"
  gBusquedas.Filtro = " and Estado = 'A'"
  gBusquedas.Orden = "cod_Plan"
  frmBusquedas.Show vbModal, Me
  If gBusquedas.Resultado <> "" Then
    txtCodigo.Text = gBusquedas.Resultado
    txtDescripcion.Text = gBusquedas.Resultado2
  End If
End If

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Consulta = "select cod_Plan,descripcion from fnd_Planes"
  gBusquedas.Filtro = " and Estado = 'A'"
  gBusquedas.Orden = "cod_Plan"
  frmBusquedas.Show vbModal, Me
  If gBusquedas.Resultado <> "" Then
    txtCodigo.Text = gBusquedas.Resultado
    txtDescripcion.Text = gBusquedas.Resultado2
  End If
End If

End Sub

Private Sub txtCodigo_LostFocus()

On Error GoTo vError

strSQL = "select descripcion from fnd_Planes where cod_Plan = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
  txtDescripcion.Text = rs!Descripcion
rs.Close

Exit Sub

vError:
    txtDescripcion.Text = ""
End Sub


