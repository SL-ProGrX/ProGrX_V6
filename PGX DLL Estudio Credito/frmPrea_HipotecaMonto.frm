VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPrea_HipotecaMonto 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Credito Hipotecario: Gastos Asociados"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      _Version        =   1572864
      _ExtentX        =   20135
      _ExtentY        =   12091
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
      ItemCount       =   3
      Item(0).Caption =   "Traspaso de bienes Inmuebles"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid(0)"
      Item(1).Caption =   "Cancelaciones de Hipoteca"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid(1)"
      Item(2).Caption =   "Constitución Hipoteca"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGrid(2)"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6615
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   11668
         _StockProps     =   64
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPrea_HipotecaMonto.frx":0000
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6615
         Index           =   1
         Left            =   -70000
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   11668
         _StockProps     =   64
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPrea_HipotecaMonto.frx":07A4
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6615
         Index           =   2
         Left            =   -70000
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   11668
         _StockProps     =   64
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPrea_HipotecaMonto.frx":0F03
         AppearanceStyle =   1
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   405
      Left            =   1920
      TabIndex        =   2
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
      TabIndex        =   3
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
Attribute VB_Name = "frmPrea_HipotecaMonto"
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

strSQL = "exec spCrdPrea_Hipotecas_Gastos '" & txtExpediente.Text & "', '" & pTipo & "'"
Call OpenRecordSet(rs, strSQL)

Select Case pTipo
    Case "CAN"
        Index = 1
    Case "CON"
        Index = 2
    Case "TRA"
        Index = 0
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
    
    If Index = 0 Then
        .Col = 6
        .Text = Format(rs!Imp_Traspaso, "Standard")
    End If
    
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

Call sbLista("TRA")


End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim pTipo As String

Select Case Item.Index
    Case 0
        pTipo = "TRA"
    Case 1
        pTipo = "CAN"
    Case 2
        pTipo = "CON"
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

Dim pBI As String, pCANH As String, pCONSH As String, pTipo As String

pBI = "Null"
pCANH = "Null"
pCONSH = "Null"

Select Case Index
    Case 0 'Traspaso
        pTipo = "BIIM"
        pBI = pParametro
        
    Case 1 'Cancelaciones
        pTipo = "CANH"
        pCANH = pParametro
        
    Case 2 'Constitucion
        pTipo = "CONH"
        pCONSH = pParametro
End Select


strSQL = "exec spCrdPreaAvaluosHipoteca '" & txtExpediente.Text & "', " & pBI & ", " & pCANH & ", " & pCONSH & ", '" & pTipo & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

'spCrdPreaAvaluosHipoteca](
'    @COD_PREANALISIS varchar(20)
'    ,@ID_PARAM_BI INT NULL
'    ,@ID_PARAM_CANH INT NULL
'    ,@ID_PARAM_CONSH INT NULL
'    ,@TIPO_PARAM CHAR(4)
'    ,@USUARIO VARCHAR(50)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub
