VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFNDCDPS_Tasas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estructura de Tasas para CDP's"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10965
   ScaleWidth      =   11370
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4935
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   11415
      _Version        =   1572864
      _ExtentX        =   20135
      _ExtentY        =   8705
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
      Item(0).Caption =   "Mantenimiento"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "lswTasas"
      Item(0).Control(2)=   "txtFiltro(0)"
      Item(0).Control(3)=   "cboVencimiento"
      Item(0).Control(4)=   "cboCupon"
      Item(0).Control(5)=   "txtTasa"
      Item(0).Control(6)=   "Label2(3)"
      Item(0).Control(7)=   "Label2(2)"
      Item(0).Control(8)=   "Label2(0)"
      Item(0).Control(9)=   "ShortcutCaption1"
      Item(1).Caption =   "Activación"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "chkActivo"
      Item(1).Control(1)=   "txtNotas"
      Item(1).Control(2)=   "btnActiva_Aplica"
      Item(2).Caption =   "Bitácora"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswBitacora"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3975
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   5295
         _Version        =   1572864
         _ExtentX        =   9340
         _ExtentY        =   7011
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
         Appearance      =   21
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswTasas 
         Height          =   3255
         Left            =   5520
         TabIndex        =   5
         Top             =   1560
         Width           =   5775
         _Version        =   1572864
         _ExtentX        =   10186
         _ExtentY        =   5741
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
         Appearance      =   21
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswBitacora 
         Height          =   4335
         Left            =   -69880
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   11175
         _Version        =   1572864
         _ExtentX        =   19711
         _ExtentY        =   7646
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
         Appearance      =   21
      End
      Begin XtremeSuiteControls.PushButton btnActiva_Aplica 
         Height          =   375
         Left            =   -63040
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmFNDCDPS_Tasas.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   375
         Left            =   -68080
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Poner la configuración de Tasas para CDPs Activa?"
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
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   5295
         _Version        =   1572864
         _ExtentX        =   9340
         _ExtentY        =   635
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
      Begin XtremeSuiteControls.ComboBox cboVencimiento 
         Height          =   330
         Left            =   5520
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.ComboBox cboCupon 
         Height          =   330
         Left            =   7920
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   330
         Left            =   10320
         TabIndex        =   9
         ToolTipText     =   "Presione F4"
         Top             =   1200
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1575
         Left            =   -68080
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   6735
         _Version        =   1572864
         _ExtentX        =   11880
         _ExtentY        =   2778
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   480
         Width           =   5895
         _Version        =   1572864
         _ExtentX        =   10398
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Tasas por Vencimiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   12
         Top             =   960
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Vencimiento"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   10440
         TabIndex        =   11
         Top             =   960
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tasa"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   7920
         TabIndex        =   10
         Top             =   960
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cupones"
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
         WordWrap        =   -1  'True
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4215
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   8895
      _Version        =   524288
      _ExtentX        =   15690
      _ExtentY        =   7435
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
      MaxCols         =   493
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDCDPS_Tasas.frx":0727
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración de Tasas para Certificados a Plazo"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   6972
   End
   Begin XtremeShortcutBar.ShortcutCaption scNivel 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   11415
      _Version        =   1572864
      _ExtentX        =   20135
      _ExtentY        =   661
      _StockProps     =   14
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
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmFNDCDPS_Tasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub sbPlanes_Load()

On Error GoTo vError

If vPaso Or scNivel.Tag = "" Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

vPaso = True

lsw.ListItems.Clear

txtFiltro(0).Text = fxSysCleanTxtInject(txtFiltro(0).Text)

'Planes
strSQL = "select Pl.cod_Operadora, Pl.cod_Plan,Pl.Descripcion,Asg.registro_Fecha,ASg.Registro_Usuario" _
       & " from Fnd_Planes Pl left join FND_CDPS_TASA_PLANES Asg on Pl.cod_operadora = Asg.cod_Operadora" _
       & " and Pl.cod_Plan = Asg.Cod_Plan and Asg.COD_TASA_REF = '" & scNivel.Tag _
       & "' where Pl.Estado = 'A' and Pl.Tipo_CDP = 1 and Pl.PAGO_CUPONES = 1" _
       & " and (Pl.Cod_Plan like '%" & txtFiltro(0).Text & "%' or Pl.Descripcion like '%" & txtFiltro(0).Text & "%')" _
       & " order by isnull(Asg.Cod_Plan,'ZZZZZZZZZZZZ') asc,Pl.cod_Plan asc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!COD_PLAN)
     itmX.Tag = rs!COD_OPERADORA
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!REGISTRO_USUARIO & ""
     itmX.SubItems(3) = rs!REGISTRO_FECHA & ""
     
     If Not IsNull(rs!REGISTRO_FECHA) Then itmX.Checked = True
 rs.MoveNext
Loop
rs.Close


vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbTasas_Add()

On Error GoTo vError

If vPaso Or scNivel.Tag = "" Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

vPaso = True
'spFnd_CDP_Tasas_Add(@TasaCod varchar(10), @fCuponId int, @PlazoId int, @Tasa dec(7,2), @Estado smallint = 1, @Usuario varchar(30))

strSQL = "exec spFnd_CDP_Tasas_Add '" & scNivel.Tag & "', " & cboCupon.ItemData(cboCupon.ListIndex) _
       & ", " & cboVencimiento.ItemData(cboVencimiento.ListIndex) & ", " & CCur(txtTasa.Text) & ", 1, '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Pass = 0 Then
    Me.MousePointer = vbDefault
    MsgBox rs!Mensaje, vbExclamation
End If

vPaso = False

Me.MousePointer = vbDefault
Call sbTasas_Load

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbTasas_Load()

On Error GoTo vError

If vPaso Or scNivel.Tag = "" Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

vPaso = True

lswTasas.ListItems.Clear

strSQL = "select T.*, C.CUPON , V.PLAZO " _
       & "  from FND_CDP_TASACUPONES T" _
       & "         inner join FND_CDP_FRECUENCIACUPONES C on T.ID_FRECUENCIACUPON = C.ID_FRECUENCIACUPON" _
       & "         inner join FND_CDP_PLAZOS V on T.ID_PLAZOCUPON = V.ID_PLAZO" _
       & " Where T.COD_TASA_REF = '" & scNivel.Tag & "' and V.ID_PLAZO = " & cboVencimiento.ItemData(cboVencimiento.ListIndex)

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswTasas.ListItems.Add(, , rs!ID_TASA)
     itmX.SubItems(1) = rs!Plazo
     itmX.SubItems(2) = rs!Cupon
     itmX.SubItems(3) = Format(rs!Tasa, "Standard")
'     itmX.SubItems(2) = rs!REGISTRO_USUARIO & ""
'     itmX.SubItems(3) = rs!REGISTRO_FECHA & ""

 rs.MoveNext
Loop
rs.Close


vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbDetalle_Consulta()
 
tcMain.Item(0).Selected = True
 
 Call sbPlanes_Load
 Call sbTasas_Load
End Sub

Private Sub btnActiva_Aplica_Click()

If vPaso Or scNivel.Tag = "" Then Exit Sub

If Len(txtNotas.Text) < 10 Then
    MsgBox "Indique una nota válida!", vbExclamation
    Exit Sub
End If

On Error GoTo vError

strSQL = "exec spFnd_CDP_Tasas_Activacion '" & scNivel.Tag & "', '" & txtNotas.Text & "', '" & glogon.Usuario & "', " & chkActivo.Value
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
   MsgBox "Proceso de Activación/Inactivación Realizado!", vbInformation
   
    strSQL = "select * from FND_CDPS_TASA_REF" _
          & " order by COD_TASA_REF"
    Call sbCargaGridLocal(vGrid, 5, strSQL)
  
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboVencimiento_Click()
If vPaso Then Exit Sub

Call sbTasas_Load

End Sub

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

With lswBitacora.ColumnHeaders
   .Clear
   .Add , , "Id", 1200
   .Add , , "Código", 1200, vbCenter
   .Add , , "Descripción", 3000
   .Add , , "Usuario", 1800, vbCenter
   .Add , , "Fecha", 1800
   .Add , , "Movimiento", 1800, vbCenter
   .Add , , "V. Anterior", 1800, vbRightJustify
   .Add , , "V.Actual", 1800, vbRightJustify
   .Add , , "Notas", 3800
End With


'B.COD_TASA_REF , B.Movimiento, B.NOTAS, B.V_ANTERIOR, B.V_ACTUAL, B.REGISTRO_FECHA, B.REGISTRO_USUARIO
'    , C.DESCRIPCION as 'MODELO_DESC'
With lsw.ColumnHeaders
   .Clear
   .Add , , "Planes", 1500
   .Add , , "Descripción", 3500
   .Add , , "Usuario", 10
   .Add , , "Fecha", 10
End With
lsw.Checkboxes = True

With lswTasas.ColumnHeaders
   .Clear
   .Add , , "Id", 1000
   .Add , , "Vencimiento", 1500
   .Add , , "Cupón", 1500
   .Add , , "Tasa", 1000, vbRightJustify
End With

strSQL = "select ID_FRECUENCIACUPON AS 'IdX', Cupon as 'ItmX'" _
       & " From FND_CDP_FRECUENCIACUPONES" _
       & " Where Estado = 1" _
       & " order by FRECUENCIA_MESES"
Call sbCbo_Llena_New(cboCupon, strSQL, False, True)

strSQL = "select ID_PLAZO as 'IdX', Plazo as 'ItmX'" _
       & " From FND_CDP_PLAZOS" _
       & " Where Estado = 1" _
       & " Order by PLAZO_MESES"
Call sbCbo_Llena_New(cboVencimiento, strSQL, False, True)

txtTasa.Text = "0.00"

scNivel.Tag = ""
scNivel.Caption = "- Seleccione un Modelo de Tasa- "

strSQL = "select * from FND_CDPS_TASA_REF" _
      & " order by COD_TASA_REF"
Call sbCargaGridLocal(vGrid, 5, strSQL)


Call Formularios(Me)
Call RefrescaTags(Me)

lsw.Enabled = vGrid.Enabled
lswTasas.Enabled = vGrid.Enabled

End Sub

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1 'Codigo
        vGrid.Text = CStr(rs!COD_TASA_REF)
     Case 2 'descripcion
        vGrid.Text = CStr(rs!Descripcion)
     Case 3 'Divisa
        vGrid.Text = CStr(rs!cod_Divisa)
     Case 4 ' Activo
        vGrid.Value = rs!Activo
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub



Private Function fxGuardar() As Long
Dim vCuenta As String, vCuentaSalida As String

On Error GoTo vError
        
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

fxGuardar = 0
If Trim(vGrid.Text) = "" Then Exit Function

Dim pCodigo As String, pDescripcion As String, pDivisa As String, pActivo As Integer


vGrid.Col = 1
pCodigo = Trim(vGrid.Text)
vGrid.Col = 2
pDescripcion = Trim(vGrid.Text)
vGrid.Col = 3
pDivisa = Trim(vGrid.Text)
vGrid.Col = 4
pActivo = vGrid.Value

strSQL = "exec spFnd_CDP_Tasa_Config_Add '" & pCodigo & "','" & pDescripcion & "','" & pDivisa & "', " & pActivo & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
  vGrid.Col = 1
  Call Bitacora(rs!Movimiento, "CDPS Modelo de Tasas:   " & vGrid.Text)
Else
  MsgBox rs!Mensaje, vbCritical
End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vCodigo As String

If vPaso Or scNivel.Tag = "" Then Exit Sub

On Error GoTo vError

vCodigo = scNivel.Tag

If Item.Checked Then
   strSQL = "exec spFnd_CDP_Tasa_Plan_Add '" & scNivel.Tag & "', '" & Item.Text & "', '" & glogon.Usuario & "', 'A'"
   Item.SubItems(2) = glogon.Usuario
   Item.SubItems(3) = Date

Else
   strSQL = "exec spFnd_CDP_Tasa_Plan_Add '" & scNivel.Tag & "', '" & Item.Text & "', '" & glogon.Usuario & "', 'E'"
   
   Item.SubItems(2) = ""
   Item.SubItems(3) = ""
   
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswTasas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError

If vPaso Or scNivel.Tag = "" Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

vPaso = True

cboVencimiento.Text = Item.SubItems(1)
cboCupon.Text = Item.SubItems(2)
txtTasa.Text = Format(Item.SubItems(3), "Standard")


vPaso = False

Me.MousePointer = vbDefault
Call sbTasas_Load

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbModelo_Bitacora()

On Error GoTo vError

If vPaso Or scNivel.Tag = "" Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "exec spFnd_CDP_Tasa_Bitacora_Consulta '" & scNivel.Tag & "'"

lswBitacora.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswBitacora.ListItems.Add(, , rs!ID_TASA_CAMBIO)
     itmX.SubItems(1) = rs!COD_TASA_REF
     itmX.SubItems(2) = rs!MODELO_DESC
     itmX.SubItems(3) = rs!REGISTRO_USUARIO
     itmX.SubItems(4) = rs!REGISTRO_FECHA
     itmX.SubItems(5) = rs!Movimiento
     itmX.SubItems(6) = Format(rs!V_ANTERIOR, "Standard")
     itmX.SubItems(7) = Format(rs!V_ACTUAL, "Standard")
     itmX.SubItems(8) = rs!Notas
     
 rs.MoveNext
Loop
rs.Close


vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 2 Then
    Call sbModelo_Bitacora
End If

End Sub

Private Sub txtFiltro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
  If Index = 0 Then
     Call sbPlanes_Load
  Else
     Call sbTasas_Load
  End If
End If

End Sub

Private Sub txtTasa_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
 Call sbTasas_Add
End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub
If Col <> 5 Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1
scNivel.Tag = vGrid.Text
vGrid.Col = 2
scNivel.Caption = vGrid.Text
vGrid.Col = 4
chkActivo.Value = vGrid.Value

txtNotas.Text = ""

Call sbDetalle_Consulta

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol >= (vGrid.MaxCols - 1) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
'        strSQL = "exec spFndSeguridad_ApAnul_Delete '" & vGrid.Text & "', '" & glogon.Usuario & "'"
'        Call ConectionExecute(strSQL)
'
'        strSQL = vGrid.Text
'        vGrid.col = 1
'        Call Bitacora("Elimina", "CDPS Modelo de Tasas:   " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

        scNivel.Tag = ""
        scNivel.Caption = ""
        lsw.ListItems.Clear
        lswTasas.ListItems.Clear

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




