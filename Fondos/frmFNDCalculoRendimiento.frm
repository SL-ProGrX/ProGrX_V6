VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDCalculoRendimiento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cálculo de Rendimientos"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTCP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   4
      Top             =   6255
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtTasa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8520
      TabIndex        =   7
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vhGrid 
      Height          =   3372
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   9852
      _Version        =   524288
      _ExtentX        =   17378
      _ExtentY        =   5948
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
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDCalculoRendimiento.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   732
      Left            =   7680
      TabIndex        =   13
      Top             =   5160
      Width           =   2292
      _Version        =   1572864
      _ExtentX        =   4043
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Aplicar Rendimientos"
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
      Picture         =   "frmFNDCalculoRendimiento.frx":0674
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   372
      Left            =   1080
      TabIndex        =   14
      Top             =   5520
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   1920
      TabIndex        =   15
      Top             =   120
      Width           =   6492
      _Version        =   1572864
      _ExtentX        =   11456
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1920
      TabIndex        =   16
      Top             =   480
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3240
      TabIndex        =   17
      Top             =   480
      Width           =   5172
      _Version        =   1572864
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "TCP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   2400
      TabIndex        =   12
      Top             =   4920
      Width           =   1332
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "TBP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   4920
      Width           =   1332
   End
   Begin VB.Label Label7 
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   852
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   10332
   End
   Begin VB.Label Label3 
      Caption         =   "Tasa anual ..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   852
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Historial de Tasas Aplicadas para Cálculo de Rendimientos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   5175
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmFNDCalculoRendimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean


Private Sub sbProcesoRendimiento(vOperadora As Long, vPlan As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vRendimientos As Currency, vCasos As Long, vPendientes As Long

On Error GoTo vError

Me.MousePointer = vbHourglass


lbl.Caption = "Revisando Tasas [Espere]..."
lbl.Refresh

strSQL = "exec spFndAjusteTasaCntVencidos " & vOperadora & ", '" & vPlan & "'"
Call ConectionExecute(strSQL)

lbl.Caption = "Inicializando [Espere]..."
lbl.Refresh

'Paso 1: Inicializa
strSQL = "exec spFndRndGenPlanMain " & vOperadora & ",'" & vPlan & "','" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular _
       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'," & CCur(txtTasa) & ",1,'" & App.ProductName & "'," & CCur(txtTCP.Text)
Call OpenRecordSet(rs, strSQL)
  vRendimientos = rs!Rendimiento
  vCasos = rs!Casos
  vPendientes = rs!Casos - rs!Procesados
rs.Close

PrgBar.Visible = True
PrgBar.Max = vCasos + 1
PrgBar.Value = 1


'Paso 2: Procesa Casos de 100 en 100
Do While vPendientes > 0
    
    lbl.Caption = "Procesando Registro: " & PrgBar.Value & " de " & PrgBar.Max _
                & vbCrLf & vbCrLf & " [Rendimiento Aplicado: " & Format(vRendimientos, "Standard") & "]"
    lbl.Refresh
    
    strSQL = "exec spFndRndGenPlanMain " & vOperadora & ",'" & vPlan & "','" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular _
           & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'," & CCur(txtTasa) & ",2,'" & App.ProductName & "'," & CCur(txtTCP.Text)
    Call OpenRecordSet(rs, strSQL)
      vRendimientos = rs!Rendimiento
      vCasos = rs!Casos
      vPendientes = rs!Pendientes
      If PrgBar.Value < PrgBar.Max Then PrgBar.Value = rs!Procesados
    rs.Close

Loop

'Paso 3: Cierra Proceso y Asiento
strSQL = "exec spFndRndGenPlanMain " & vOperadora & ",'" & vPlan & "','" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular _
       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'," & CCur(txtTasa) & ",3,'" & App.ProductName & "'," & CCur(txtTCP.Text)
Call OpenRecordSet(rs, strSQL)
  vRendimientos = rs!Rendimiento
  vCasos = rs!Casos
rs.Close

MsgBox "Rendimientos distribuidos satisfactoriamente!" & vbCrLf & vbCrLf _
     & " --> Rendimiento: " & Format(vRendimientos, "Standard") & vbCrLf _
     & " --> Casos      : " & Format(vCasos, "###,###,##0"), vbInformation


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboOperadora_Click()
Call txtCodigo_LostFocus
End Sub

Private Sub cboOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub


Private Sub CmdAplicar_Click()
Dim curPorcentaje As Currency, vOperadora As Long
Dim vPlan As String, rs As New ADODB.Recordset
Dim vFecha As Date

vFecha = fxFechaServidor

If Trim(txtCodigo) = "" Or txtTasa = "" Or txtTasa = "0" Or txtTasa = "." Or txtTCP.Text = "" Or txtTCP.Text = "0" Then
   
   MsgBox "Faltan Datos", vbExclamation

Else
   
   If Format(vFecha, "yyyymmdd") < Format(dtpCorte.Value, "yyyymmdd") Then
     MsgBox "La fecha de Corte es Mayor a la Fecha del Sistema...", vbExclamation
     Exit Sub
   End If
   
   vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
   vPlan = Trim(txtCodigo)
      
   rs.Source = "select isnull(count(*),0) as Existe" _
             & " from fnd_planes " _
             & " where cod_operadora = " & vOperadora & " and cod_plan = '" & vPlan & "' and rend_corte <= '" _
             & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
   rs.Open , glogon.Conection, adOpenStatic
   If rs!Existe = 0 Then
      MsgBox "La fecha de Corte es Menor al Ultimo Corte Realizado para Este Plan...", vbExclamation
      Exit Sub
   End If
   rs.Close
      
   Me.MousePointer = vbHourglass
    Call sbProcesoRendimiento(vOperadora, vPlan)
    Call Bitacora("Genera", "Rendimiento Ope:" & vOperadora & " Plan:" & vPlan)
   
   
   Me.MousePointer = vbDefault
   
  
   txtCodigo_LostFocus
   
   
End If

End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and calcula_rend = 1"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!COD_PLAN
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

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion
End Sub

Private Sub Form_Load()
Dim strSQL As String


vModulo = 18 'Fondo de Inversion
vhGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

dtpCorte.Value = fxFechaServidor

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = " And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
             & " and calcula_rend = 1"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If


End Sub




Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass

txtTasa.Locked = True
txtTasa.Text = 0
dtpCorte.Enabled = False

strSQL = "Select Descripcion,rend_corte,isnull(UltTasa, Tasa_Base) as 'UltTasa',Tasa_Base,UTILIZA_TASA_FLUCTUANTE,UTILIZA_TBP" _
       & ", dbo.fxFndTasaReferencia('TBP') as 'TBP'" _
       & ", dbo.fxFndTasaReferencia('TCP') as 'TCP'" _
       & " from Fnd_Planes" _
       & " where Cod_Operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " And Cod_Plan='" & Trim(txtCodigo) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtDescripcion = Trim(rs!Descripcion)
   
   txtTasa.Enabled = True
   txtTCP.Enabled = True
   
   dtpCorte.Enabled = True
   
   txtTasa.Text = rs!UltTasa & ""
   txtTCP.Text = rs!TCP
   
   If rs!utiliza_tbp = 1 Then
      txtTasa.Locked = True
      txtTasa.Text = rs!TBP
      
      txtTCP.Locked = True
      txtTCP.Text = rs!TCP
      
   Else
      txtTasa.Locked = False
   End If
      
   If rs!UTILIZA_TASA_FLUCTUANTE = 1 Then
      txtTasa.Locked = False
      txtTCP.Locked = False
   Else
      txtTasa.Locked = True
      txtTCP.Locked = True
   End If

    strSQL = "select Top 24 corte,tasa,TCP,usuario,fecha_sys" _
           & " from FND_HISTORIAL_REND" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " And Cod_Plan = '" & Trim(txtCodigo) & "' order by IDx desc"
    Call sbCargaGrid(vhGrid, 5, strSQL)
    vhGrid.MaxRows = vhGrid.MaxRows - 1


Else
   txtCodigo = ""
   txtDescripcion = ""
   vhGrid.MaxRows = 0
End If
rs.Close

lbl.Caption = ""
PrgBar.Visible = False

Me.MousePointer = vbDefault


End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTasa.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
                & " and calcula_rend = 1"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtTasa.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub

