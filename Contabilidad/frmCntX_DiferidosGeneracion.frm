VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCntX_DiferidosGeneracion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generación de Asientos de Movimientos Diferidos"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   8280
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   10335
      _Version        =   1310723
      _ExtentX        =   18230
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGenerar 
         Height          =   495
         Left            =   8760
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Generar"
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
         Picture         =   "frmCntX_DiferidosGeneracion.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl_Main 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10335
      _Version        =   1310723
      _ExtentX        =   18224
      _ExtentY        =   9123
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
      Item(0).Caption =   "Generación"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "chkTodos"
      Item(0).Control(2)=   "scTitulo"
      Item(1).Caption =   "Resultados"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "txt"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4455
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   10335
         _Version        =   1310723
         _ExtentX        =   18230
         _ExtentY        =   7858
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
      End
      Begin XtremeSuiteControls.FlatEdit txt 
         Height          =   4815
         Left            =   -70000
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1310723
         _ExtentX        =   18230
         _ExtentY        =   8493
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
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
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   210
         Left            =   360
         TabIndex        =   6
         Top             =   460
         Width           =   210
         _Version        =   1310723
         _ExtentX        =   379
         _ExtentY        =   379
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   10335
         _Version        =   1310723
         _ExtentX        =   18230
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione los Diferidos Pendientes que desea Procesar:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Generación de Asientos Diferidos del Periodo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   8895
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "frmCntX_DiferidosGeneracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass


TabControl_Main.Item(0).Selected = True

strSQL = "exec spCntX_Diferido_Pendientes " & gCntX_Parametros.CodigoConta & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes

Call OpenRecordSet(rs, strSQL, 0)
lsw.ListItems.Clear

Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!cod_diferido)
        itmX.SubItems(1) = rs!cod_difPlantilla
        itmX.SubItems(2) = rs!Descripcion
        itmX.SubItems(3) = Format(rs!monto_diferir - rs!acumulado, "Standard")
        itmX.SubItems(4) = rs!Consecutivo
     
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub chkTodos_Click()
Dim i As Long

For i = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdGenerar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCodPlantilla As Long, vCodDiferido As Long


Me.MousePointer = vbHourglass
On Error GoTo vError


If Not fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
  Me.MousePointer = vbDefault
  MsgBox "El Periodo Actual se Encuentra Cerrado o no se ha creado, verifique...", vbExclamation
  Exit Sub
End If

TabControl_Main.Item(1).Selected = True
txt = ""

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
    DoEvents
    
    vCodPlantilla = lsw.ListItems.Item(i).Text
    vCodDiferido = lsw.ListItems.Item(i).SubItems(1)
    
    'proc spCntX_Diferido_Asiento(@Contabilidad int, @Plantilla int, @Diferido int, @Anio int, @Mes smallint, @Usuario varchar(30))
    strSQL = strSQL & Space(10) & "exec spCntX_Diferido_Asiento " & gCntX_Parametros.CodigoConta & "," & vCodPlantilla & "," & vCodDiferido _
           & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes & ",'" & glogon.Usuario & "'"
    
    txt = txt & vbCrLf & "PROCESANDO : DIF:" & vCodDiferido & " PLANTILLA: " & vCodPlantilla & "-" & lsw.ListItems.Item(i).SubItems(2)
    
    
    If Len(strSQL) > 20000 Then
       Call ConectionExecute(strSQL)
       strSQL = ""
    End If
    
  End If
  
Next i

'Procesa ultimo lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If



Me.MousePointer = vbDefault

MsgBox "Asientos de Plantillas Diferidas...Generadas Satisfactoriamente...", vbInformation
Call sbCargaLsw


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Plantilla", 1200
    .Add , , "Diferido", 1200, vbCenter
    .Add , , "Descripción", 3800
    .Add , , "Pendiente", 2100, vbRightJustify
    .Add , , "Consecutivo", 1400, vbRightJustify
End With

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Call sbCargaLsw
End Sub


