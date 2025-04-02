VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInvPaquetes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paquetes / Combos"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   9210
   Begin VB.Frame fraActivacion 
      Caption         =   "Activar Paquete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   4935
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   3240
         TabIndex        =   26
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   252641283
         CurrentDate     =   37682
      End
      Begin VB.CheckBox chkDomingos 
         Caption         =   "Domingos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox chkSabados 
         Caption         =   "Sábados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chkViernes 
         Caption         =   "Viernes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkJueves 
         Caption         =   "Jueves"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkMiercoles 
         Caption         =   "Miércoles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkMartes 
         Caption         =   "Martes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkLunes 
         Caption         =   "Lunes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   315
         Left            =   3240
         TabIndex        =   27
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   252641283
         CurrentDate     =   37682
      End
      Begin MSComCtl2.DTPicker dtpHoraInicio 
         Height          =   315
         Left            =   3240
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   252641282
         CurrentDate     =   37682
      End
      Begin MSComCtl2.DTPicker dtpHoraCorte 
         Height          =   315
         Left            =   3240
         TabIndex        =   29
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   252641282
         CurrentDate     =   37682
      End
      Begin VB.Label Label3 
         Caption         =   "Hora Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   25
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hora Inicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   24
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   1560
         X2              =   1560
         Y1              =   360
         Y2              =   1680
      End
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Width           =   6735
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5610
      Width           =   1455
   End
   Begin VB.TextBox txtImpuestos 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5250
      Width           =   1455
   End
   Begin VB.TextBox txtDescuento 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   8055
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1005
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2172
      Left            =   0
      TabIndex        =   32
      Top             =   1560
      Width           =   9132
      _Version        =   524288
      _ExtentX        =   16108
      _ExtentY        =   3831
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmInvPaquetes.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblLineas 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5520
      TabIndex        =   31
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label lblCantidad 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5520
      TabIndex        =   30
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Paquete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   10
      Top             =   5610
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "(+) Impuestos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   9
      Top             =   5250
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "(-) Descuento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmInvPaquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 32
 
 vGrid.AppearanceStyle = fxGridStyle

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = 0
txtCodigo = ""

txtDescripcion = ""
txtNotas = ""

vGrid.MaxRows = 1
vGrid.MaxCols = 7
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

txtSubTotal = 0
txtDescuento = 0
txtImpuestos = 0
txtTotal = 0

txtCodigo.Enabled = True


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpHoraInicio.Value = dtpInicio.Value
dtpHoraCorte.Value = dtpCorte.Value

chkLunes.Value = vbUnchecked
chkMartes.Value = vbUnchecked
chkMiercoles.Value = vbUnchecked
chkJueves.Value = vbUnchecked
chkViernes.Value = vbUnchecked
chkSabados.Value = vbUnchecked
chkDomingos.Value = vbUnchecked

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "descripcion"
'       gBusquedas.Orden = "descripcion"
'       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from pv_paquetes" _
       & " where cod_paquete = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_paquete
  txtCodigo = rs!cod_paquete
  
  txtDescripcion = rs!Descripcion
  txtNotas = rs!Notas & ""

  dtpInicio.Value = rs!Fecha_Inicio
  dtpCorte.Value = rs!Fecha_Corte
  dtpHoraInicio.Value = rs!frecuencia_horai
  dtpHoraCorte.Value = rs!frecuencia_horac
  
  chkLunes.Value = rs!frecuencia_lunes
  chkMartes.Value = rs!frecuencia_martes
  chkMiercoles.Value = rs!frecuencia_miercoles
  chkJueves.Value = rs!frecuencia_jueves
  chkViernes.Value = rs!frecuencia_viernes
  chkSabados.Value = rs!frecuencia_sabado
  chkDomingos.Value = rs!frecuencia_domingo
  

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,D.porc_utilidad" _
         & ",D.precio,D.imp_ventas, (D.cantidad * (D.precio + D.precio * D.porc_utilidad / 100)) " _
         & " + ((D.cantidad * (D.precio + D.precio * D.porc_utilidad / 100)) * (D.imp_ventas / 100)) as Total" _
         & " from pv_paquetes_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_paquete = " & rs!cod_paquete _
         & " order by D.Linea"
  
  Call sbCargaGrid(vGrid, 7, strSQL)
  
  Call sbCalculaTotales
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

On Error GoTo vError

vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "E", 1)


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer
Dim curCantidad As Currency
Dim rs As New ADODB.Recordset

On Error GoTo vError


If vEdita Then
    strSQL = "update pv_paquetes set descripcion = '" & UCase(txtDescripcion) _
           & "',notas = '" & txtNotas & "',user_modifica = '" & glogon.Usuario _
           & "',fecha_modifica = dbo.MyGetdate(),fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',frecuencia_horai = '" _
           & Format(dtpHoraInicio.Value, "hh:mm:ss") & "',frecuencia_horac = '" & Format(dtpHoraCorte.Value, "hh:mm:ss") _
           & "',frecuencia_lunes = " & chkLunes.Value & ",frecuencia_martes = " & chkMartes.Value _
           & ",frecuencia_miercoles = " & chkMiercoles.Value & ",frecuencia_jueves = " & chkJueves.Value _
           & ",frecuencia_viernes = " & chkViernes.Value & ",frecuencia_sabado = " & chkSabados.Value _
           & ",frecuencia_domingo = " & chkDomingos.Value _
           & " where cod_paquete = " & vCodigo
   Call ConectionExecute(strSQL)

   Call Bitacora("Modifica", "Paquete (Oferta): " & vCodigo)

Else
    strSQL = "select isnull(max(cod_paquete),0) + 1 as Paquete from pv_paquetes"
    Call OpenRecordSet(rs, strSQL)
     vCodigo = rs!Paquete
    rs.Close
    txtCodigo = vCodigo
    
    strSQL = "insert pv_paquetes(cod_paquete,descripcion,fecha_crea,user_crea,notas" _
           & ",fecha_inicio,fecha_corte,frecuencia_horai,frecuencia_horac" _
           & ",frecuencia_lunes,frecuencia_martes,frecuencia_miercoles,frecuencia_jueves" _
           & ",frecuencia_viernes,frecuencia_sabado,frecuencia_domingo)" _
           & " values(" & vCodigo & ",'" & UCase(txtDescripcion) & "',dbo.MyGetdate(),'" _
           & glogon.Usuario & "','" & txtNotas & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & Format(dtpHoraInicio.Value, "hh:mm:ss") _
           & "','" & Format(dtpHoraCorte.Value, "hh:mm:ss") & "'," & chkLunes.Value _
           & "," & chkMartes.Value & "," & chkMiercoles.Value & "," & chkJueves.Value _
           & "," & chkViernes.Value & "," & chkSabados.Value & "," & chkDomingos.Value & ")"
    Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Paquete (Oferta): " & vCodigo)
End If

txtCodigo.Enabled = True

'Guardar Detalle de la Orden
strSQL = "delete pv_paquetes_detalle" _
         & " where cod_paquete = " & vCodigo
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.col = 1
    strSQL = "insert pv_paquetes_detalle(linea,cod_paquete,cod_producto,cantidad" _
           & ",porc_utilidad,precio,imp_ventas,imp_consumo) values(" & i & "," & vCodigo & ",'" _
           & vGrid.Text & "'," & curCantidad & ","
    vGrid.col = 4
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.col = 5
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.col = 6
    strSQL = strSQL & CCur(vGrid.Text) & ",0)"
    Call ConectionExecute(strSQL)
  End If
Next i

'*********************************** fin

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
'  Call ConectionExecute(strSQL)

'  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & vParametros.CodigoEmpresa)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Integer, vSQL As String

vSQL = ""

Select Case UCase(ButtonMenu.Key)
  Case "REPBOLETA"
     
     i = MsgBox("Desea visualizar solo el paquete Actual", vbYesNo)
     If i = vbYes Then vSQL = "{PV_PAQUETES.COD_PAQUETE} = " & txtCodigo

     Call sbInvReportes("PaquetesBoleta", "Boleta de Paquetes", "", vSQL)

  Case "REPLISTADOGENERAL"
     Call sbInvReportes("PaquetesListado", "PAQUETES", "Listado General", vSQL)

End Select


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_paquete"
  gBusquedas.Orden = "cod_paquete"
  gBusquedas.Consulta = "select cod_paquete,descripcion,notas from pv_paquetes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub sbCalculaTotales()
Dim curSubTotal As Currency, curIV As Currency, curTmpUtilidad As Currency
Dim curTmpPrecio As Currency, curTmpIV As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long
Dim iLineas As Integer, curCantidad As Currency

On Error GoTo vError

'**********************************************    OJO
'Revisar esta formula por la situacion del descuento, si es antes o despues del
'impuesto de ventas, por ahora está despues del impuesto

curSubTotal = 0
curIV = 0

iLineas = 0
curCantidad = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
    curTmpCant = CCur(vGrid.Text)
    vGrid.col = 4
    curTmpUtilidad = CCur(vGrid.Text)
    vGrid.col = 5
    curTmpPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curTmpIV = CCur(vGrid.Text)
    
    curTmpPrecio = curTmpPrecio + (curTmpPrecio * curTmpUtilidad / 100)
    
    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio)
    curIV = curIV + ((curTmpCant * curTmpPrecio) * (curTmpIV / 100))
 
    curCantidad = curCantidad + curTmpCant
    iLineas = iLineas + 1
 
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")
txtImpuestos = Format(curIV, "Standard")
txtTotal = Format(curSubTotal + curIV - CCur(txtDescuento), "Standard")

lblLineas.Caption = "Líneas   : " & iLineas
lblCantidad.Caption = "Cantidad : " & Format(curCantidad, "Standard")


vError:

End Sub

Private Sub sbConsultaArticulo(fila As Long, Columna As Integer, vCriterio As String)
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

'Busquedas
'1. Por Codigo del Articulo
'2. Por Codigo de Barras
'3. Por Codigo del Fabricante
vPaso = False

strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
       & " where cod_producto = '" & vCriterio & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then vPaso = True

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_barras = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_fabricante = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  MsgBox "No se encontró el Articulo en la Base de Datos...", vbExclamation
Else
  vGrid.Row = fila
  vGrid.col = 1
  vGrid.Text = rs!Cod_Producto
  vGrid.col = 2
  vGrid.Text = rs!Descripcion
  vGrid.col = 5
  vGrid.Text = CStr(rs!costo_regular)
  vGrid.col = 6
  vGrid.Text = CStr(rs!impuesto_ventas)
End If
rs.Close


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Consulta Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
End If

'Consular Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
   frmBusquedaArticulos.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   vGrid.Text = gBusquedas.Resultado
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 7
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To 7
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 1 To 7
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  Call sbCalculaTotales
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub



Private Sub vGrid_KeyPress(KeyAscii As Integer)
Dim curCantidad As Currency, curPrecio As Currency, curIV As Currency
Dim curUtilidad As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 4, 5, 6
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 4
    curUtilidad = CCur(vGrid.Text)
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curIV = CCur(vGrid.Text)
    
    curPrecio = curPrecio + (curPrecio * curUtilidad / 100)
    
    vGrid.col = 7
    vGrid.Text = (curPrecio * curCantidad) + ((curPrecio * curCantidad) * (curIV / 100))
   
   Call sbCalculaTotales
  
  Case Else 'No Aplica
End Select
vError:
End Sub


