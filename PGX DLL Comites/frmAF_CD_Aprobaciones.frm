VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_CD_Aprobaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desembolsos a Comités y Delegados: Aprobación de Desembolsos"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   120
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox gbAprobar 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   7680
      Width           =   16095
      _Version        =   1572864
      _ExtentX        =   28390
      _ExtentY        =   2355
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   615
         Index           =   0
         Left            =   10560
         TabIndex        =   6
         Top             =   600
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Rechazar"
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
         Picture         =   "frmAF_CD_Aprobaciones.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtNota 
         Height          =   975
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   7455
         _Version        =   1572864
         _ExtentX        =   13150
         _ExtentY        =   1720
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
      Begin XtremeSuiteControls.PushButton btnResolucion 
         Height          =   615
         Index           =   1
         Left            =   12240
         TabIndex        =   7
         Top             =   600
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aprobar"
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
         Picture         =   "frmAF_CD_Aprobaciones.frx":05A4
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nota:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   16095
      _Version        =   524288
      _ExtentX        =   28390
      _ExtentY        =   10821
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
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_CD_Aprobaciones.frx":0CCB
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   330
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   6975
      _Version        =   1572864
      _ExtentX        =   12303
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.CheckBox chkTodas 
      Height          =   210
      Left            =   720
      TabIndex        =   8
      Top             =   1200
      Width           =   210
      _Version        =   1572864
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Banco:"
      ForeColor       =   16777215
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   16575
   End
End
Attribute VB_Name = "frmAF_CD_Aprobaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, i As Long

Private Sub sbCasos_Load()

If vPaso Then Exit Sub

If cboBanco.ListCount <= 0 Or cboBanco.Text = "" Then
  vGrid.MaxRows = 0
  Exit Sub
End If

If cboBanco.Text = "TODOS" Then
    strSQL = "select distinct 0 as ValorX,C.noperacion,C.cod_comite,U.descripcion as Comite,C.cedula" _
           & ",S.nombre,C.Cuenta, coalesce(sum(M.monto),0) as Total" _
           & " from afi_cd_cuentas C left join Uprogramatica U on C.cod_comite = U.codigo" _
           & " left join Socios S on C.cedula = S.cedula " _
           & " left join afi_cd_cuentas_actividades M on C.nOperacion = M.nOperacion" _
           & " where C.id_banco = " & cboBanco.ItemData(cboBanco.ListIndex) & " and C.estado = 'S'" _
           & " group by C.noperacion,C.cod_comite,U.descripcion,C.cedula" _
           & ",S.nombre,C.Cuenta"
    
Else
    strSQL = "select distinct 0 as ValorX,C.noperacion,C.cod_comite,U.descripcion as Comite,C.cedula" _
           & ",S.nombre,C.Cuenta, coalesce(sum(M.monto),0) as Total" _
           & " from afi_cd_cuentas C left join Uprogramatica U on C.cod_comite = U.codigo" _
           & " left join Socios S on C.cedula = S.cedula " _
           & " left join afi_cd_cuentas_actividades M on C.nOperacion = M.nOperacion" _
           & " where C.estado = 'S'" _
           & " group by C.noperacion,C.cod_comite,U.descripcion,C.cedula" _
           & ",S.nombre,C.Cuenta"
End If

Call sbCargaGrid(vGrid, 8, strSQL)

'Elimina Linea en Blanco
vGrid.MaxRows = vGrid.MaxRows - 1

End Sub

Private Sub sbAprobacion()

If vGrid.MaxRows = 0 Then
  MsgBox "No hay informacion para procesar", vbInformation, "Información"
  Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

For i = 1 To vGrid.MaxRows
   vGrid.Row = i
       vGrid.Col = 1
       If vGrid.Value = vbChecked Then
           vGrid.Col = 2
           
           'Activa y Registra Asiento
           strSQL = "exec spAFI_CD_AsientoCuentas '" & vGrid.Text & "', '" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular & "'"
           
           Call ConectionExecute(strSQL)
                      
       End If
Next i

Me.MousePointer = vbDefault
MsgBox "Aprobación Realizada", vbInformation, "Información"
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
 
End Sub

Private Sub sbRechazar()

If vGrid.MaxRows = 0 Then
  MsgBox "No hay informacion para procesar", vbInformation, "Información"
  Exit Sub
End If

On Error GoTo vError
Me.MousePointer = vbHourglass


For i = 1 To vGrid.MaxRows
   vGrid.Row = i
       vGrid.Col = 1
       If vGrid.Value = vbChecked Then
           vGrid.Col = 2
           strSQL = "update afi_cd_cuentas set estado = 'R' " _
                  & "where noperacion = '" & vGrid.Text & "'"
           Call ConectionExecute(strSQL)
       End If
Next i

Me.MousePointer = vbDefault
MsgBox "Se rechaza la Operación", vbInformation, "Información"



Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub btnResolucion_Click(Index As Integer)
Select Case Index
  Case 1 'APROBAR"
    Call sbAprobacion

  Case 0 'Rechazado"
    Call sbRechazar

End Select

Call sbCasos_Load
   
End Sub

Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount <= 0 Then
  vGrid.MaxRows = 0
  Exit Sub
End If

  Call sbCasos_Load
End Sub

Private Sub chkTodas_Click()
For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 1
    vGrid.Value = chkTodas.Value
Next i
End Sub

Private Sub Form_Load()
vModulo = 40

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

strSQL = "select B.id_banco as 'IdX', B.descripcion as 'itmX'" _
       & " from Tes_bancos B inner join afi_cd_cuentas C on B.id_banco = C.id_banco" _
       & " Where C.Estado = 'S'" _
       & " group by B.id_Banco, B.descripcion "
       
vPaso = True
Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
vPaso = False

End Sub
