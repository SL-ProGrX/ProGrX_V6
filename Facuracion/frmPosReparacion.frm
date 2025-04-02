VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPosReparacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Articulos en Reparación"
   ClientHeight    =   6156
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11928
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6156
   ScaleWidth      =   11928
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacion.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacion.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacion.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacion.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacion.frx":19AB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Width           =   5175
   End
   Begin VB.TextBox txtDocumento 
      Appearance      =   0  'Flat
      DataField       =   "e"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9600
      TabIndex        =   9
      ToolTipText     =   "Codigo Proveedor"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      DataField       =   "e"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      ToolTipText     =   "Codigo Proveedor"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   10335
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11928
      _ExtentX        =   21040
      _ExtentY        =   635
      BandCount       =   2
      _CBWidth        =   11928
      _CBHeight       =   360
      _Version        =   "6.7.9816"
      Child1          =   "tlb"
      MinHeight1      =   264
      Width1          =   3864
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   312
      Width2          =   3816
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   312
         Left            =   4020
         TabIndex        =   16
         Top             =   24
         Width           =   7836
         _ExtentX        =   13822
         _ExtentY        =   550
         ButtonWidth     =   2265
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Traslado"
               Key             =   "Traslado"
               Object.ToolTipText     =   "Traslado a Taller"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Información"
               Key             =   "Info"
               Object.ToolTipText     =   "Info. de Seguimiento"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Entrega"
               Key             =   "Entrega"
               Object.ToolTipText     =   "Entrega a Cliente"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   132
         TabIndex        =   2
         Top             =   48
         Width           =   3708
         _ExtentX        =   6541
         _ExtentY        =   466
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
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   9600
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   550
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   198574083
      CurrentDate     =   37791
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   14
      Top             =   5844
      Width           =   11928
      _ExtentX        =   21040
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Estado General de la Orden"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3952
            MinWidth        =   3952
            Object.ToolTipText     =   "Usuario Ingresa"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
            Object.ToolTipText     =   "Fecha Ingreso"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   480
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   0
      TabIndex        =   17
      Top             =   2280
      Width           =   11772
      _Version        =   524288
      _ExtentX        =   20765
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosReparacion.frx":2EC22
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "# Documento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8520
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   11760
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8520
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "# Orden"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmPosReparacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vMascara As String
Dim vScroll As Boolean, vUltimo(2) As String

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_orden from POS_REPARACION"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_orden > '" & txtCodigo & "' order by cod_orden asc"
    Else
       strSQL = strSQL & " where cod_orden < '" & txtCodigo & "' order by cod_orden desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_orden)
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
vModulo = 33
End Sub

Private Sub Form_Load()

On Error GoTo vError


 vModulo = 33
 vGrid.AppearanceStyle = fxGridStyle

 vMascara = "0000000000"
 vEdita = True
 
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
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

vCodigo = ""
txtCodigo = ""

txtCedula = ""
txtNombre = ""

dtpFecha.Value = fxFechaServidor
txtFecha = Format(dtpFecha.Value, "yyyy/mm/dd hh:mm:ss")
dtpFecha.Visible = True

txtDocumento = ""
txtNotas = ""

vGrid.MaxRows = 1
vGrid.MaxCols = 6
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

StatusBarX.Panels.Item(1) = "Solicitado"
StatusBarX.Panels.Item(2) = ""
StatusBarX.Panels.Item(3) = ""


End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      Call sbToolBar(tlb, "edicion")
      txtCedula.SetFocus
      
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNotas.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
      Call sbBorrar
    
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
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

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select O.*,C.Nombre" _
       & " from POS_REPARACION O inner join pv_clientes C on O.cedula = C.cedula" _
       & " where O.cod_orden = '" & Format(xCodigo, vMascara) & "'"
       

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = Format(xCodigo, vMascara)
  txtCodigo = Format(xCodigo, vMascara)
  
  txtCedula = rs!Cedula
  txtNombre = rs!Nombre
  
  txtFecha = Format(rs!genera_fecha, "yyyy/mm/dd hh:mm:ss")
  dtpFecha.Value = rs!genera_fecha
  
  txtDocumento = rs!Documento
  txtNotas = rs!nota & ""
  
  Select Case rs!Estado
    Case "SC" 'Solicitada
       StatusBarX.Panels(1) = "Solicitada"
    Case "TC" 'Trasladada
       StatusBarX.Panels(1) = "Trasladada"
    Case "TP" 'Trasladada Parcial
       StatusBarX.Panels(1) = "Trasladada Parcial"
    Case "RC" 'Recibido de Taller
       StatusBarX.Panels(1) = "Recibido de Taller"
    Case "RP" 'Recibido de Taller
       StatusBarX.Panels(1) = "Recibido de Taller Parcial"
    Case "EC" 'Entregada
       StatusBarX.Panels(1) = "Entregado al Cliente"
    Case "EP" 'Entregada Parcial Cliente
       StatusBarX.Panels(1) = "Entrega Parcial al Cliente"
  End Select
  
       StatusBarX.Panels(2) = rs!genera_user & ""
       StatusBarX.Panels(3) = rs!genera_fecha & ""
  

  strSQL = "select D.*,P.descripcion as Producto, X.descripcion as Proveedor" _
         & " from POS_REPARACION_DETALLE D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " inner join cxp_proveedores X on D.cod_proveedor = X.cod_proveedor" _
         & " where D.cod_orden = '" & vCodigo & "' order by D.Linea"
  Call sbCargaGridLocal(vGrid, 6, strSQL)
  
  vGrid.Enabled = True
  
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


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows

  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!cod_producto)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        Select Case rs!Estado
           Case "S"
             vGrid.CellNote = "Estado : Solicitado." & vbCrLf _
                  & "Ingreso : " & rs!fecha_entrada & ""
           Case "T"
             vGrid.CellNote = "Estado : Trasladado." & vbCrLf _
                  & "Traslado : " & rs!fecha_traslado & ""
           Case "R"
             vGrid.CellNote = "Estado : Recibido." & vbCrLf _
                  & "Recibo Taller : " & rs!FECHA_RECIBO & ""
           Case "E"
             vGrid.CellNote = "Estado : Entregado." & vbCrLf _
                  & "Entrega Cliente : " & rs!FECHA_Entrega & ""

        End Select
        vGrid.TextTip = TextTipFixed
     Case 2
        vGrid.Text = CStr(rs!producto)
     Case 3
        vGrid.Text = CStr(rs!nserie)
     Case 4
        vGrid.Text = CStr(rs!cod_Factura)
     
     Case 5
        vGrid.Text = CStr(rs!cod_proveedor)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = CStr(rs!Proveedor)
        vGrid.TextTip = TextTipFixed
     Case 6
        vGrid.Text = CStr(rs!Detalle)
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

On Error GoTo vError

If txtCedula = "" Then vMensaje = vMensaje & vbCrLf & " - No se especifico cliente..."
If txtNotas = "" Then vMensaje = vMensaje & vbCrLf & " - Indique las notas sobre la orden y reparacion"

vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbGuardar()
Dim strSQL As String, i As Integer, rs As New ADODB.Recordset
Dim vFecha As Date

On Error GoTo vError

If dtpFecha.Visible Then
  vFecha = dtpFecha.Value
Else
  vFecha = fxFechaServidor
End If


'If vEdita And Mid(StatusBarX.Panels(1).Text, 1, 1) <> "S" Then
'   MsgBox "No se Puede Editar una Boleta que ha sido trasladada...", vbExclamation
'   Exit Sub
'End If

'' Por el momento no se permite modificar porque afecta inventarios directamente. A menos
'' que se ponga un proceso de Autorizacion
If vEdita Then
   MsgBox "No se Puede Editar una Boleta que ha sido trasladada...", vbExclamation
   Exit Sub
End If


If Not vEdita Then
    'Consecutivo de la Orden
    strSQL = "select isnull(max(cod_orden),0) + 1 as Ultimo from POS_REPARACION"
    Call OpenRecordSet(rs, strSQL)
      vCodigo = Format(rs!ultimo, vMascara)
    rs.Close
    txtCodigo = vCodigo
    
    strSQL = "insert POS_REPARACION(estado,proceso,cod_orden,cedula,documento,genera_user," _
           & "genera_fecha,nota) values('SC','P','" & txtCodigo & "','" & txtCedula & "','" _
           & txtDocumento & "','" & glogon.Usuario & "',dbo.MyGetdate(),'" & txtNotas & "')"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Registra", "Orden de Servicio Reparacion: " & vCodigo)

Else

    strSQL = "update pos_reparacion set cedula = '" & txtCedula & "',documento = '" & txtDocumento _
           & "',genera_user = '" & glogon.Usuario & "',nota = '" & txtNotas _
           & "' where cod_orden = '" & vCodigo & "'"
    Call ConectionExecute(strSQL)

    Call Bitacora("Modifica", "Orden de Servicio Reparacion: " & vCodigo)

End If



txtCodigo.Enabled = True

'Guardar Detalle de la Orden
strSQL = "delete POS_REPARACION_DETALLE" _
         & " where cod_orden = '" & txtCodigo & "'"
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  
  If vGrid.Text <> "" Then
    
    vGrid.col = 1
    strSQL = "insert POS_REPARACION_DETALLE(linea,cod_orden,cod_producto,nserie,cod_factura,cod_proveedor,detalle,estado)" _
           & " values(" & i & ",'" & vCodigo & "','" & vGrid.Text & "','"
    vGrid.col = 3
    strSQL = strSQL & vGrid.Text & "','"
    vGrid.col = 4
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 5
    strSQL = strSQL & vGrid.Text & ",'"
    vGrid.col = 6
    strSQL = strSQL & vGrid.Text & "','S')"
    
    Call ConectionExecute(strSQL)
  
  End If
Next i

'Crear Movimientos de Entrada
Call sbPosSRAfectaInv(vCodigo, "S", 0, 0)

Call sbConsulta(vCodigo)
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

     vSQL = "{POS_REPARACION.cod_orden} = '" & txtCodigo & "'"
     Call sbPosReportesSR("BoletaEntrada", "Servicio de Reparación", "Recepción", vSQL)

 
End Sub


Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)

GLOBALES.gTag = txtCodigo.Text

Select Case Button.Key
  Case "Traslado"
     'Call sbTraslado
     frmPosReparacionTraslado.Show vbModal
  Case "Info"
     frmPosReparacionInfo.Show vbModal
  Case "Entrega"
     frmPosReparacionEntregaCliente.Show vbModal
End Select

Call sbConsulta(vCodigo)

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCedula_LostFocus()
txtNombre = fxSIFCCodigos("D", txtCedula, "clientes")
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_orden"
  gBusquedas.Orden = "cod_orden"
  gBusquedas.Consulta = "select O.cod_orden,C.nombre" _
          & " from POS_REPARACION O inner join pv_clientes C on O.cedula = C.cedula"
  gBusquedas.Filtro = ""
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
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
  vGrid.Text = rs!cod_producto
  vGrid.col = 2
  vGrid.Text = rs!Descripcion
End If
rs.Close


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer

'Abrir Nueva Linea / Y conserval ultimos valores
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.col = 4
    vUltimo(0) = vGrid.Text
    
    vGrid.col = 5
    vUltimo(1) = vGrid.Text
    vUltimo(2) = vGrid.CellNote
    
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    
    vGrid.col = 4
    vGrid.Text = CStr(vUltimo(0))
    
    vGrid.col = 5
    vGrid.Text = CStr(vUltimo(1))
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = CStr(vUltimo(2))
    vGrid.TextTip = TextTipFixed
    
    
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


'Buscar Proveedor
If vGrid.ActiveCol = 5 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = fxSIFCCodigos("D", vGrid.Text, "Proveedores")
  vGrid.TextTip = TextTipFixed
End If

'Consular Proveedor
If vGrid.ActiveCol = 5 And KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 5
   vGrid.Text = gBusquedas.Resultado
   vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
   vGrid.CellNote = gBusquedas.Resultado2
   vGrid.TextTip = TextTipFixed
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.MaxCols
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    
    vGrid.col = 4
    vUltimo(0) = vGrid.Text
    
    vGrid.col = 5
    vUltimo(1) = vGrid.Text
    vUltimo(2) = vGrid.CellNote

    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow

    vGrid.col = 4
    vGrid.Text = CStr(vUltimo(0))
    
    vGrid.col = 5
    vGrid.Text = CStr(vUltimo(1))
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = CStr(vUltimo(2))
    vGrid.TextTip = TextTipFixed

End If

End Sub




