VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_Notica_No_Cotizantes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificaciones para No Cotizantes"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboRango 
      Height          =   330
      Left            =   7200
      TabIndex        =   4
      Top             =   840
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6800
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   14295
      _Version        =   1441793
      _ExtentX        =   25215
      _ExtentY        =   1296
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmAF_Notica_No_Cotizantes.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   1
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar"
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
         Picture         =   "frmAF_Notica_No_Cotizantes.frx":0700
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   2
         Left            =   11280
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Enviar Notificación"
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
         Picture         =   "frmAF_Notica_No_Cotizantes.frx":086A
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboAviso 
         Height          =   330
         Left            =   9120
         TabIndex        =   10
         Top             =   240
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   3
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmAF_Notica_No_Cotizantes.frx":0F83
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   4
         Left            =   7440
         TabIndex        =   15
         ToolTipText     =   "Actualizar Estadisticas"
         Top             =   240
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
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
         Picture         =   "frmAF_Notica_No_Cotizantes.frx":168A
      End
   End
   Begin XtremeSuiteControls.ComboBox cboInforme 
      Height          =   330
      Left            =   11280
      TabIndex        =   12
      Top             =   840
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   0
      TabIndex        =   13
      Top             =   2160
      Width           =   13095
      _Version        =   524288
      _ExtentX        =   23098
      _ExtentY        =   8705
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
      MaxCols         =   29
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Notica_No_Cotizantes.frx":1D92
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   14895
      _Version        =   1441793
      _ExtentX        =   26273
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   3
      Top             =   840
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Rango"
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
      Index           =   7
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inst/Empr."
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notificación para No Cotizantes"
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
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   765
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmAF_Notica_No_Cotizantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean


Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 26
    vHeaders.Headers(1) = "Notifica?"
    vHeaders.Headers(2) = "Cédula"
    vHeaders.Headers(3) = "Nombre"
    vHeaders.Headers(4) = "F.Ingreso"
    vHeaders.Headers(5) = "Estado"
    vHeaders.Headers(6) = "Email"
    vHeaders.Headers(7) = "Tel.Celular"
    vHeaders.Headers(8) = "Tel.Habitación"
    vHeaders.Headers(9) = "Tel.Trabajo"
    vHeaders.Headers(10) = "F.Ult.Apo.Obrero"
    vHeaders.Headers(11) = "F.Ult.Apo.Patronal"
    vHeaders.Headers(12) = "Dias Obrero"
    vHeaders.Headers(13) = "Dias Patronal"
    vHeaders.Headers(14) = "F.Activa"
    vHeaders.Headers(15) = "F.Suspende"
    vHeaders.Headers(16) = "Empresa/Inst."
    vHeaders.Headers(17) = "UP"
    vHeaders.Headers(18) = "UP Descripción"
    vHeaders.Headers(19) = "Promotor"
    vHeaders.Headers(20) = "Aporte Obrero"
    vHeaders.Headers(21) = "Capitalización"
    vHeaders.Headers(22) = "Aporte Patronal"
    vHeaders.Headers(23) = "Fondos Acumulados"
    vHeaders.Headers(24) = "Saldo de Creditos"
    vHeaders.Headers(25) = "Morosidad"
    
    vHeaders.Headers(26) = "Provincia"
    vHeaders.Headers(27) = "Cantón"
    vHeaders.Headers(28) = "Distrito"
    vHeaders.Headers(29) = "Dirección"

 If cboInforme.ItemData(cboInforme.ListIndex) <> 1 Then
    Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Control_NoCotiza_" + cboInforme.Text)
 Else
    Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Control_NoCotiza_" + cboRango.Text)
 End If
 
End Sub

Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pInstitucion As String

If cboInstitucion.Text = "TODOS" Then
    pInstitucion = "Null"
Else
    pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If


 If cboInforme.ItemData(cboInforme.ListIndex) <> 1 Then
    strSQL = "exec spPAT_AsociadosSinAportes_Consulta " & cboInforme.ItemData(cboInforme.ListIndex) & ", 0, " & pInstitucion
 Else
    strSQL = "exec spPAT_AsociadosSinAportes_Consulta " & cboInforme.ItemData(cboInforme.ListIndex) & ", " & cboRango.ItemData(cboRango.ListIndex) & ", " & pInstitucion
 End If

Call OpenRecordSet(rs, strSQL)



With vGrid
  .MaxRows = 0
  Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Value = chkTodas.Value
     .Col = 2
     .Text = Trim(rs!Cedula)
     .Col = 3
     .Value = Trim(rs!Nombre)
     .Col = 4
     .Text = Format(rs!FechaIngreso, "yyyy-mm-dd")
     .Col = 5
     .Text = Trim(rs!Estado_Asociado)
     .Col = 6
     .Text = Trim(rs!Email)
     .Col = 7
     .Text = Trim(rs!Telefono_Celular & "")
     .Col = 8
     .Text = Trim(rs!Telefono_Habitacion & "")
     .Col = 9
     .Text = Trim(rs!Telefono_Trabajo)
     
     .Col = 10
     .Text = Format(rs!UltimoAporteObrero, "yyyy-mm-dd")
     .Col = 11
     .Text = Format(rs!UltimoAportePatronal, "yyyy-mm-dd")
     .Col = 12
     .Text = CStr(rs!Dias_Aporte_Obrero & "")
     .Col = 13
     .Text = CStr(rs!Dias_Aporte_Patronal & "")
     
     .Col = 14
     .Text = Format(rs!Fecha_Activo & "", "yyyy-mm-dd")
     .Col = 15
     .Text = Format(rs!Fecha_Suspendido & "", "yyyy-mm-dd")
     
     .Col = 16
     .Text = Trim(rs!Institucion & "")
     .Col = 17
     .Text = Trim(rs!UP & "")
     .Col = 18
     .Text = Trim(rs!UP_Desc & "")
     .Col = 19
     .Text = Trim(rs!Promotor_Desc)
     
     .Col = 20
     .Text = Format(rs!Aporte_Obrero, "Standard")
     .Col = 21
     .Text = Format(rs!Capitalización, "Standard")
     .Col = 22
     .Text = Format(rs!Aporte_Patronal, "Standard")
     .Col = 23
     .Text = Format(rs!Fondos_Acumulado, "Standard")
     .Col = 24
     .Text = Format(rs!Creditos_Saldo, "Standard")
     .Col = 25
     .Text = Format(rs!Morosidad, "Standard")
     
     
     .Col = 26
     .Text = Trim(rs!Provincia_Desc & "")
     .Col = 27
     .Text = Trim(rs!Canton_Desc & "")
     .Col = 28
     .Text = Trim(rs!Distrito_Desc & "")
     .Col = 29
     .Text = Trim(rs!Direccion & "")
     

   rs.MoveNext
  Loop
  rs.Close
End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbNotificacion()
Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""

ProgressBarX.Visible = True

With vGrid
    ProgressBarX.Max = .MaxRows
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
           .Col = 2
            strSQL = strSQL & Space(10) & "exec spPAT_AsociadosSinAportes_Notifica '" & .Text & "', " & cboAviso.ItemData(cboAviso.ListIndex) _
                   & "  , '" & glogon.Usuario & "', "
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
        
        ProgressBarX.Value = i
    Next i

    'Lote Final
    If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If

End With

ProgressBarX.Visible = False

Me.MousePointer = vbDefault
MsgBox "Notificaciones: " & cboAviso.Text & ", enviadas satisfactoriamente!", vbInformation


Exit Sub

vError:
  Me.MousePointer = vbDefault
  ProgressBarX.Visible = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbInforme()

MsgBox "Pendiente Revision con Usuario", vbInformation

End Sub


Private Sub sbEstadistica_Update()
On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPAT_AsociadosSinAportes_RecalculaFechas"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Actualización de Fechas de Pago Valido, actualizadas satisfactoriamente!", vbInformation


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnAccion_Click(Index As Integer)
Select Case Index
    Case 0 'Buscar
        Call sbBuscar
        
    Case 1 'Exportar
        Call sbExportar
        
    Case 2 'Notificacion
        Call sbNotificacion
    
    Case 3 'Informe
        Call sbInforme
    
    Case 4 'Estadistica de Fecha
        Call sbEstadistica_Update
    

End Select
End Sub


Private Sub cboInforme_Click()
If cboInforme.ItemData(cboInforme.ListIndex) = 1 Then
   cboRango.Enabled = True
Else
   cboRango.Enabled = False
End If
End Sub

Private Sub chkTodas_Click()
Dim i As Long

With vGrid
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        .Value = chkTodas.Value
    Next i
End With

End Sub

Private Sub Form_Load()
vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle
vGrid.MaxRows = 0


strSQL = "select cod_Institucion as 'IdX', Descripcion as 'ItmX'" _
       & " from Instituciones WHERE ACTIVA = 1"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

strSQL = "select Linea_Id as 'IdX', Descripcion as 'ItmX' " _
       & "  From AFI_SOCIOS_SIN_APORTES_RANGOS Where Activo = 1 order by Dia_Desde"
Call sbCbo_Llena_New(cboRango, strSQL, False, True)
       
cboInforme.AddItem "Activos"
cboInforme.ItemData(cboInforme.ListCount - 1) = CStr(1)
cboInforme.AddItem "Suspendidos"
cboInforme.ItemData(cboInforme.ListCount - 1) = CStr(0)
cboInforme.AddItem "Condición Espcial"
cboInforme.ItemData(cboInforme.ListCount - 1) = CStr(2)
cboInforme.AddItem "Suspend. + C.E."
cboInforme.ItemData(cboInforme.ListCount - 1) = CStr(3)
cboInforme.Text = "Activos"
       
cboAviso.AddItem "1er Aviso"
cboAviso.ItemData(cboAviso.ListCount - 1) = CStr(1)
cboAviso.AddItem "2do Aviso"
cboAviso.ItemData(cboAviso.ListCount - 1) = CStr(2)
cboAviso.AddItem "3er Aviso"
cboAviso.ItemData(cboAviso.ListCount - 1) = CStr(3)
cboAviso.AddItem "Notificación"
cboAviso.ItemData(cboAviso.ListCount - 1) = CStr(4)

cboAviso.Text = "1er Aviso"

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
gbAccion.Width = Me.Width - 250
vGrid.Width = Me.Width - 320

vGrid.Height = Me.Height - (vGrid.Top + 700)
ProgressBarX.Width = vGrid.Width

End Sub
