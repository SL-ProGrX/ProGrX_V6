VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPrea_PrendaMonto 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Montos de Traspaso y Constitución"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7455
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   13095
      _Version        =   1572864
      _ExtentX        =   23098
      _ExtentY        =   13150
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Traspaso de Bienes Muebles"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid(0)"
      Item(1).Caption =   "Constitución Prenda"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid(1)"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6975
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   12855
         _Version        =   524288
         _ExtentX        =   22675
         _ExtentY        =   12303
         _StockProps     =   64
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
         MaxCols         =   6
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "frmPrea_PrendaMonto.frx":0000
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6975
         Index           =   1
         Left            =   -70000
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   12855
         _Version        =   524288
         _ExtentX        =   22675
         _ExtentY        =   12303
         _StockProps     =   64
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
         MaxCols         =   6
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "frmPrea_PrendaMonto.frx":084A
         AppearanceStyle =   1
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   405
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   706
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Expediente"
      BackColor       =   -2147483633
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
End
Attribute VB_Name = "frmPrea_PrendaMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, i As Long

Private Sub sbLista(pTipo As String)
Dim Index As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCrdPrea_Prendas_Gastos '" & txtExpediente.Text & "', '" & pTipo & "'"
Call OpenRecordSet(rs, strSQL)

Select Case pTipo
    Case "T"
        Index = 0
    Case "C"
        Index = 1
End Select


vPaso = True

With vGrid(Index)
  .MaxRows = 0
  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    .Col = 1
    .CellTag = rs!ID_PARAM
    If rs!Asigna = 1 Then
        .Value = rs!Asigna
    End If
    .Col = 2
    .Text = Format(rs!Monto_Min, "Standard")
    .Col = 3
    .Text = Format(rs!Monto_Max, "Standard")
    .Col = 4
    .Text = Format(rs!Gastos, "Standard")
    .Col = 5
    .Text = Format(rs!Honorarios, "Standard")
    
    .Col = 6
    .Text = Format(rs!Total, "Standard")
    
    rs.MoveNext
  Loop
  rs.Close
  
End With

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()

txtExpediente.Text = GLOBALES.gTag
tcMain.Item(0).Selected = True

Call sbLista("T")


End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim pTipo As String

Select Case Item.Index
    Case 0
        pTipo = "T"
    Case 1
        pTipo = "C"
End Select

Call sbLista(pTipo)

End Sub


Private Sub vGrid_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim pParametro As Long, pValor As Integer

If vPaso Then Exit Sub


With vGrid(Index)
  .Row = Row
  .Col = 1
  pValor = .Value
  pParametro = .CellTag
  
  'Desmarca todas las filas menos la actual
  vPaso = True
    For i = 1 To .MaxRows
       .Row = i
       If i <> Row Then
          .Col = 1
          If .Value = vbChecked Then
          .Value = vbUnchecked
          End If
       End If
    Next i
  vPaso = False

End With



On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pTipo As String


Select Case Index
    Case 0 'Traspaso
        pTipo = "T"
        
    Case 1 'Constitucion
        pTipo = "C"
End Select


strSQL = "exec spCRD_PreaAsignaHonorariosPren '" & txtExpediente.Text & "', " & pParametro & ", '" & pTipo & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

'spCRD_PreaAsignaHonorariosPren]
'(
'    @pPreanalisis VARCHAR(20),
'    @pIdParam INT NULL,
'    @pProceso CHAR(1),
'    @pUsuario VARCHAR(30)
    
Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

