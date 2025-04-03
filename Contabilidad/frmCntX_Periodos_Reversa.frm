VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCntX_Periodos_Reversa 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reversión de Periodos Contables"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   14235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswBitacora 
      Height          =   4695
      Left            =   7080
      TabIndex        =   2
      Top             =   1680
      Width           =   7125
      _Version        =   1441793
      _ExtentX        =   12559
      _ExtentY        =   8281
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   7125
      _Version        =   1441793
      _ExtentX        =   12559
      _ExtentY        =   8281
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   14055
      _Version        =   1441793
      _ExtentX        =   24791
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Reversión"
      ForeColor       =   4210752
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReversa 
         Height          =   615
         Left            =   12120
         TabIndex        =   4
         Top             =   600
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Reversar"
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
         Picture         =   "frmCntX_Periodos_Reversa.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   855
         Left            =   6960
         TabIndex        =   5
         Top             =   360
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   1508
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblCierre 
         Height          =   255
         Left            =   12120
         TabIndex        =   7
         ToolTipText     =   "Corte"
         Top             =   360
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "2024-01-01"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   855
         Left            =   6240
         TabIndex        =   6
         Top             =   360
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   1508
         _StockProps     =   14
         Caption         =   "Notas"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   7095
      _Version        =   1441793
      _ExtentX        =   12515
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Cierres Registrados"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scCierre 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   1320
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Bitacora de Cierres"
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
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   14415
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Reversion de Cierres"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
Attribute VB_Name = "frmCntX_Periodos_Reversa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbCortes()

On Error GoTo vError

lblCierre.Caption = ""
lsw.ListItems.Clear
lswBitacora.ListItems.Clear

strSQL = "select TOP 60 PERIODO_CORTE , ESTADO, CIERRE_FECHA, CIERRE_USUARIO" _
       & " FROM CNTX_PERIODOS WHERE COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " AND ESTADO = 'C' ORDER BY PERIODO_CORTE DESC"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Format(rs!PERIODO_CORTE, "yyyy-mm-dd"))
     itmX.SubItems(1) = rs!CIERRE_FECHA & ""
     itmX.SubItems(2) = rs!CIERRE_USUARIO & ""
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:

End Sub


Private Sub sbCortes_Bitacora(pCierre As String)

On Error GoTo vError

lblCierre.Caption = pCierre
lswBitacora.ListItems.Clear

strSQL = "select P.PERIODO_CORTE as 'CORTE' " _
       & ", L.MOVIMIENTO, L.REGISTRO_FECHA, L.REGISTRO_USUARIO" _
       & " FROM CNTX_PERIODOS_LOG L" _
       & "  INNER JOIN CNTX_PERIODOS P ON L.COD_CONTABILIDAD = P.COD_CONTABILIDAD AND L.ANIO = P.ANIO AND L.MES = P.MES" _
       & " WHERE L.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & "   AND L.ANIO = YEAR('" & pCierre & "') AND L.MES = MONTH('" & pCierre _
       & "') ORDER BY L.REGISTRO_FECHA DESC"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswBitacora.ListItems.Add(, , Format(rs!Corte, "yyyy-mm-dd"))
     itmX.SubItems(1) = IIf(rs!Movimiento = "C", "Cierra", "Reversa")
     itmX.SubItems(2) = rs!REGISTRO_FECHA & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:

End Sub

Private Sub btnReversa_Click()

On Error GoTo vError

Dim vMensaje As String

vMensaje = ""

If lblCierre.Caption = "" Then
  vMensaje = vMensaje & vbCrLf & " - No se ha indicado ningún Periodo"
End If

txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)

If Len(txtNotas.Text) < 10 Then
  vMensaje = vMensaje & vbCrLf & " - Especifique una nota válida para la reversión"
End If

If Len(vMensaje) > 0 Then
    MsgBox vMensaje, vbExclamation
    Exit Sub
End If
     
Dim i As Integer
i = MsgBox("Esta seguro que desea >> Reversar << este Periodo: " & lblCierre.Caption, vbYesNo)
If i = vbNo Then
    Exit Sub
End If

Me.MousePointer = vbHourglass

Dim pAnio As Integer, pMes As Integer, pCierre As Date

pCierre = lblCierre.Caption
pAnio = Year(pCierre)
pMes = Month(pCierre)


strSQL = "exec spCntX_Periodo_Reversa " & gCntX_Parametros.CodigoConta _
       & ", " & pAnio & ", " & pMes & ", '" & txtNotas.Text & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Reversa", "Periodo Contable: " & lblCierre.Caption)

Call sbCortes

Me.MousePointer = vbDefault

MsgBox "Cierre Reversado Satisfactoriamente...", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
  vModulo = 20
   
 Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

 With lsw.ColumnHeaders
    .Add , , "Corte", 2000
    .Add , , "Fecha", 2500
    .Add , , "Usuario", 2500
 End With
 
 With lswBitacora.ColumnHeaders
    .Add , , "Corte", 2000
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Fecha", 2500
    .Add , , "Usuario", 2500
 End With
  
 Call sbCortes
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Call sbCortes_Bitacora(Item.Text)
End Sub



