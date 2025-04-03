VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCC_App_Log 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estadisticas de Consumo del App"
   ClientHeight    =   8745
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7212
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   13812
      _Version        =   1310723
      _ExtentX        =   24363
      _ExtentY        =   12721
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
      Item(0).Caption =   "Estadística"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Detalle"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "Label2(1)"
      Item(1).Control(1)=   "lblItem"
      Item(1).Control(2)=   "vGridDet"
      Item(2).Caption =   "Analisis de Ingreso"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "rbIngreso(0)"
      Item(2).Control(1)=   "Label3"
      Item(2).Control(2)=   "rbIngreso(1)"
      Item(2).Control(3)=   "vGridPin"
      Begin XtremeSuiteControls.RadioButton rbIngreso 
         Height          =   372
         Index           =   0
         Left            =   -66280
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   3132
         _Version        =   1310723
         _ExtentX        =   5524
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "con ingreso a la App"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6372
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   11052
         _Version        =   524288
         _ExtentX        =   19494
         _ExtentY        =   11239
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
         MaxCols         =   491
         ScrollBars      =   2
         SpreadDesigner  =   "frmCC_App_Log.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridDet 
         Height          =   6132
         Left            =   -68800
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   11412
         _Version        =   524288
         _ExtentX        =   20130
         _ExtentY        =   10816
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
         MaxCols         =   490
         ScrollBars      =   2
         SpreadDesigner  =   "frmCC_App_Log.frx":0716
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.RadioButton rbIngreso 
         Height          =   372
         Index           =   1
         Left            =   -63040
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   3132
         _Version        =   1310723
         _ExtentX        =   5524
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "sin Ingreso a la App"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGridPin 
         Height          =   6012
         Left            =   -70000
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   13692
         _Version        =   524288
         _ExtentX        =   24151
         _ExtentY        =   10604
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
         MaxCols         =   490
         ScrollBars      =   2
         SpreadDesigner  =   "frmCC_App_Log.frx":0DB1
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   372
         Left            =   -69880
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   3252
         _Version        =   1310723
         _ExtentX        =   5736
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Personas con registro Nuevo de PIN "
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblItem 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -68800
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   8532
      End
      Begin VB.Label Label2 
         Caption         =   "Item:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -69520
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   732
      End
   End
   Begin XtremeSuiteControls.PushButton Btn_Buscar 
      Height          =   492
      Left            =   4680
      TabIndex        =   0
      Top             =   960
      Width           =   1572
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Consultar"
      BackColor       =   -2147483633
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
      Picture         =   "frmCC_App_Log.frx":1476
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1680
      TabIndex        =   8
      Top             =   960
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   550
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
      Height          =   312
      Left            =   3120
      TabIndex        =   9
      Top             =   960
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   550
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "APPs: Estadísticas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   4452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rango"
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
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   732
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCC_App_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Public Sub sbCargaGrid_Local(vGrid As Object, vGridMaxCol As Integer, strSQL As String, pColIni As Integer)
Dim rs As New ADODB.Recordset, i As Integer
Dim pCon As New ADODB.Connection, pStrCon As String
Dim pUser As String, pKey As String

Me.MousePointer = vbHourglass

On Error GoTo vError

pUser = glogon.Core_User
pKey = glogon.Core_Key

pStrCon = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & glogon.Servidor _
       & ";Database=PGX_BASE;APP=PGX_CORE;tcp:" & glogon.Servidor _
       & "," & SIFGlobal.PuertosDisponibles & ";"

With pCon
  .CommandTimeout = 15
  .Mode = adModeReadWrite
  .CursorLocation = adUseClient
  
  .Open pStrCon, pUser, pKey
  .CommandTimeout = 360
End With

vPaso = True

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 0 'Reset

rs.Open strSQL, pCon, adOpenStatic

Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  For i = pColIni To vGrid.MaxCols
    vGrid.col = i
    vGrid.Text = RTrim(CStr(rs.Fields(i - pColIni).Value & ""))
  Next i
  rs.MoveNext
Loop
rs.Close

vPaso = False

vError:
    Me.MousePointer = vbDefault
End Sub


Private Sub Btn_Buscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

vPaso = True

strSQL = "exec spAPP_Estadistica " & gPortal.Empresa_Id & ",'" & Format(dtpInicio.Value, "yyyy/mm/dd") _
        & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

Call sbCargaGrid_Local(vGrid, 6, strSQL, 2)

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()


vGrid.AppearanceStyle = fxGridStyle
vGridDet.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -1, dtpCorte.Value)


tcMain.Item(0).Selected = True
vGrid.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbDetalle(pCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(1).Selected = True

strSQL = "exec spAPP_Estadistica_Detalle " & gPortal.Empresa_Id & ",'" & pCodigo & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
        & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

Call sbCargaGrid_Local(vGridDet, 5, strSQL, 1)


Me.MousePointer = vbDefault
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbAnalisis_Ingreso()
Dim strSQL As String
Dim i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

If rbIngreso.Item(0).Value = True Then
  i = 1
Else
  i = 0
End If

strSQL = "exec spAPP_Estadistica_Analisis " & gPortal.Empresa_Id & ",'" & Format(dtpInicio.Value, "yyyy/mm/dd") _
        & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'," & i

Call sbCargaGrid_Local(vGridPin, 6, strSQL, 1)


Me.MousePointer = vbDefault
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub rbIngreso_Click(Index As Integer)
    Call sbAnalisis_Ingreso
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 2 Then
    Call sbAnalisis_Ingreso
End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.col = 2
lblItem.Tag = vGrid.Text
vGrid.col = 3
lblItem.Caption = vGrid.Text

Call sbDetalle(lblItem.Tag)

End Sub

