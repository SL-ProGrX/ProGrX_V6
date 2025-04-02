VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmFNDPlanillaBitacora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bitácora de Planilla (Cobros) de Fondos"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3852
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   9372
      _Version        =   1441793
      _ExtentX        =   16531
      _ExtentY        =   6794
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
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3120
      TabIndex        =   0
      Top             =   2760
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
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
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
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
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   312
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
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
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   330
      Left            =   1800
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   9372
      _Version        =   1441793
      _ExtentX        =   16531
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Bitácora"
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitácora de aplicación de Aportaciones por Cobros de planillas directas"
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
      Height          =   732
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   7332
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   372
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   1332
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
      Height          =   372
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFNDPlanillaBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vScroll As Boolean



Private Sub cboInstitucion_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear

strSQL = "select * from fnd_prm_bitacora where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and proceso = " & txtProceso.Text & " Order by id_seq"
Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_seq)
     itmX.SubItems(1) = IIf(rs!Gestion = "R", "Recepción", "Envio")
     itmX.SubItems(2) = fxSIFPlanillaTipoTransac(rs!transaccion)
     itmX.SubItems(3) = rs!Documento
     itmX.SubItems(4) = rs!Usuario
     itmX.SubItems(5) = rs!fecha
     itmX.SubItems(6) = Format(rs!Monto, "Standard")
     itmX.SubItems(7) = rs!Casos
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 lsw.ListItems.Clear
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub FlatScrollBar_Change()
Dim vFecha As Currency

On Error GoTo vError

vFecha = txtProceso.Text


If vScroll Then
    
    If FlatScrollBar.Value = 1 Then
       vFecha = fxFechaProcesoSiguiente(vFecha)
    Else
       vFecha = fxFechaProcesoAnterior(vFecha)
    End If
    
    txtProceso.Text = vFecha
      
    Call cboInstitucion_Click
End If



vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strPrimero As String

vPaso = True


Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 900
    .Add , , "Gestión", 1100, vbCenter
    .Add , , "Transacción", 2000
    .Add , , "Documento", 1800
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 1800
    .Add , , "Monto", 1400, vbRightJustify
    .Add , , "Casos", 1000, vbCenter
End With

txtProceso.Text = GLOBALES.glngFechaCR

strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX from instituciones order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

strSQL = "select cod_operadora as IdX, descripcion as ItmX from fnd_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

strSQL = "select cod_plan as IdX, descripcion as ItmX from fnd_planes" _
       & " where deduce_independiente = 1 and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
Call sbCbo_Llena_New(cboPlan, strSQL, False, True)


vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

vPaso = False

Call cboInstitucion_Click

End Sub




Private Function fxSIFPlanillaTipoTransac(pTransaccion As String) As String
Dim vResultado As String
 
Select Case Trim(pTransaccion)
Case "01"
   vResultado = "Cambia Fecha de Proceso"
Case "02"
   vResultado = "Genera deducciones"
Case "03"
   vResultado = "Carga deducciones"
Case "04"
   vResultado = "Desglosa deducciones"
Case Else
   vResultado = "No.Identificado"
End Select
fxSIFPlanillaTipoTransac = vResultado
 
End Function

