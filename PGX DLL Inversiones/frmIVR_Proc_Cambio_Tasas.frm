VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmIVR_Proc_Cambio_Tasas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Cambio de Tasas"
   ClientHeight    =   7410
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2175
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   11175
      _Version        =   1441793
      _ExtentX        =   19711
      _ExtentY        =   3836
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
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   3720
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox gbFondos 
      Height          =   1212
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   11052
      _Version        =   1441793
      _ExtentX        =   19494
      _ExtentY        =   2138
      _StockProps     =   79
      Caption         =   "Fondos de Inversion:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   2
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasaAnterior 
         Height          =   315
         Left            =   3840
         TabIndex        =   19
         ToolTipText     =   "Tasa Anterior"
         Top             =   600
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUltimoCupon 
         Height          =   315
         Left            =   8160
         TabIndex        =   20
         ToolTipText     =   "Tasa Anterior"
         Top             =   120
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUltimoCorte 
         Height          =   315
         Left            =   8160
         TabIndex        =   21
         ToolTipText     =   "Tasa Anterior"
         Top             =   600
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   18
         Top             =   600
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ultimo Corte Intereses"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   6000
         TabIndex        =   17
         Top             =   120
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ultimo Cupón"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tasa"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   40
         Left            =   480
         TabIndex        =   3
         Top             =   120
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha Inicio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   615
      Left            =   8880
      TabIndex        =   5
      Top             =   6600
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmIVR_Proc_Cambio_Tasas.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtInversionId 
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   120
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtInstrumento 
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   840
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtAdministrador 
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   1200
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtPortafolio 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   1560
      Width           =   9255
      _Version        =   1441793
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeShortcutBar.ShortcutCaption scGestion 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   3840
      Width           =   11175
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Historial de Cambios: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption scGestion 
      Height          =   372
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   11172
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Gestion: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Inversión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Portafolio"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Administrador"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Instrumento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmIVR_Proc_Cambio_Tasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean
Dim itmX As ListViewItem, vFecha As Date

Private Sub sbHistorial()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from IVR_TITULOS_TASAS " _
       & " Where TITULO_ID = " & txtInversionId.Text _
       & " Order by Fecha_Inicio desc"

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Format(rs!Fecha_Inicio, "yyyy-mm-dd"))
     itmX.SubItems(1) = Format(rs!Tasa, "###,##0.0000")
     itmX.SubItems(2) = rs!Registro_Usuario & ""
     itmX.SubItems(3) = rs!Registro_Fecha & ""
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
End Sub

Private Sub btnAplicar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spIVR_TITULO_TASA_CAMBIO " & txtInversionId.Text & ", " & CCur(txtTasa.Text) _
       & " , '" & Format(dtpInicio.Value, "yyyy-MM-dd") & "', '" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Cambio de Tasa aplicada satisfactoriamente!", vbInformation

Call sbConsulta(txtInversionId.Text)
Call sbHistorial


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

On Error GoTo vError


txtInversionId.Text = gIVR_Transito.TituloId

With lsw.ColumnHeaders
    .Clear
    .Add , , "Inicio", 1800
    .Add , , "Tasa", 1200, vbRightJustify
    .Add , , "Rg.Usuario", 2800, vbCenter
    .Add , , "Rg.Fecha", 2800
End With


strSQL = "select  isnull(max(CORTE), dbo.mygetdate())  as 'CORTE'" _
       & "  From IVR_CIERRES"
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!Corte
rs.Close

Exit Sub

vError:


End Sub



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


On Error GoTo vError

Call sbConsulta(gIVR_Transito.TituloId)
Call sbHistorial

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsulta(pTituloId As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * " _
       & ", dbo.fxIVR_Titulo_Cupon_Ultimo(TITULO_ID) AS 'CUPON_CORTE' " _
       & " from vIVR_INVERSIONES" _
       & " Where Titulo_ID = " & pTituloId
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    txtInversionId.Text = rs!TITULO_ID
    
    txtTasa.Text = Format(rs!Tasa_Actual, "##0.0000")
    txtTasaAnterior.Text = Format(rs!Tasa_Actual, "##0.0000")
    
    dtpInicio.Value = rs!Cupon_Corte
    dtpInicio.MinDate = rs!Cupon_Corte
    
    txtUltimoCupon.Text = Format(rs!Cupon_Corte, "yyyy-mm-dd")
    txtUltimoCorte.Text = Format(vFecha, "yyyy-mm-dd")
    
    
    txtAdministrador.Text = rs!Administrador_Desc
    txtAdministrador.Tag = rs!Cod_Administrador
    
    txtInstrumento.Text = rs!Instrumento_Desc
    txtInstrumento.Tag = rs!Cod_Instrumento
    
    txtPortafolio.Text = rs!Portafolio_Desc
    txtPortafolio.Tag = rs!Cod_Portafolio
    
    
Else
  Me.MousePointer = vbDefault
  MsgBox "No se Localizó el registro!", vbExclamation
End If
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

