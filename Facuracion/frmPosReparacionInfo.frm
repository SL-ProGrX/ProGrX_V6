VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPosReparacionInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Boletas : Información de Seguimiento"
   ClientHeight    =   6108
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPosReparacionInfo.frx":0000
   ScaleHeight     =   6108
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMovimiento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   3495
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15896
      _ExtentY        =   7006
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Movimiento"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   5010
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Boleta Inv"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Tipo"
         Object.Width           =   1658
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   708
      Left            =   5520
      TabIndex        =   2
      Top             =   5160
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   1249
      ButtonWidth     =   2985
      ButtonHeight    =   1249
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Boleta Servicio"
            Key             =   "servicio"
            Object.ToolTipText     =   "imprimir la Boleta de Servicio"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Boleta de Inventario"
            Key             =   "inventario"
            Object.ToolTipText     =   "Boleta de Envio"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   5280
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacionInfo.frx":169B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacionInfo.frx":1D214
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Movimiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Información de Seguimiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   6252
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   9120
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmPosReparacionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBoleta As String

Private Sub sbInfo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select 'E' as Tipo,Boleta_Entrada as Boleta,Usuario_Entrada as Genera_User,Min(Fecha_Entrada) as Genera_Fecha" _
       & " From pos_reparacion_detalle" _
       & " where cod_orden = '" & GLOBALES.gTag & "' group by Boleta_Entrada,Usuario_Entrada"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , "Solicitud de Cliente")
      itmX.SubItems(1) = rs!genera_user
      itmX.SubItems(2) = rs!genera_fecha
      itmX.SubItems(3) = rs!Boleta
      itmX.SubItems(4) = rs!Tipo
 rs.MoveNext
Loop
rs.Close

strSQL = "select 'S' as Tipo,Boleta_Traslado as Boleta,Usuario_Traslado as Genera_User,Min(Fecha_Traslado) as Genera_Fecha" _
       & " From pos_reparacion_detalle" _
       & " where cod_orden = '" & GLOBALES.gTag & "' group by Boleta_Traslado,Usuario_Traslado"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  If IsNull(rs!genera_user) Then Exit Do
  Set itmX = lsw.ListItems.Add(, , "Traslado a Taller")
      itmX.SubItems(1) = rs!genera_user
      itmX.SubItems(2) = rs!genera_fecha
      itmX.SubItems(3) = rs!Boleta
      itmX.SubItems(4) = rs!Tipo
 rs.MoveNext
Loop
rs.Close

strSQL = "select 'E' as Tipo,Boleta_Ingreso as Boleta,Usuario_Recibo as Genera_User,Min(Fecha_Recibo) as Genera_Fecha" _
       & " From pos_reparacion_detalle" _
       & " where cod_orden = '" & GLOBALES.gTag & "' group by Boleta_Ingreso,Usuario_Recibo"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  If IsNull(rs!genera_user) Then Exit Do
  Set itmX = lsw.ListItems.Add(, , "Recibo de Taller")
      itmX.SubItems(1) = rs!genera_user
      itmX.SubItems(2) = rs!genera_fecha
      itmX.SubItems(3) = rs!Boleta
      itmX.SubItems(4) = rs!Tipo
 rs.MoveNext
Loop
rs.Close

strSQL = "select 'S' as Tipo,Boleta_entrega as Boleta,Usuario_Entrega as Genera_User,Min(Fecha_Entrega) as Genera_Fecha" _
       & " From pos_reparacion_detalle" _
       & " where cod_orden = '" & GLOBALES.gTag & "' group by Boleta_entrega,Usuario_Entrega"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  If IsNull(rs!genera_user) Then Exit Do
  Set itmX = lsw.ListItems.Add(, , "Entrega a Cliente")
      itmX.SubItems(1) = rs!genera_user
      itmX.SubItems(2) = rs!genera_fecha
      itmX.SubItems(3) = rs!Boleta
      itmX.SubItems(4) = rs!Tipo
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()
 
vModulo = 33
 
Call sbInfo

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lsw_Click()
 txtMovimiento.Tag = ""
 txtMovimiento.Text = ""
 

If lsw.ListItems.Count = 0 Then Exit Sub

txtMovimiento.Text = lsw.SelectedItem.Text & " [" & lsw.SelectedItem.SubItems(3) & "]"
txtMovimiento.Tag = lsw.SelectedItem.SubItems(4)
vBoleta = lsw.SelectedItem.SubItems(3)
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vSQL As String, vSubTitulo As String, vOrden As String

vSQL = ""
vOrden = ""

Select Case Button.Key
   Case "servicio"
       Select Case Mid(txtMovimiento.Text, 1, 1)
         Case "S"
           vSQL = "{POS_REPARACION.COD_ORDEN} = '" & GLOBALES.gTag & "'"
           Call sbPosReportesSR("BoletaEntrada", "SERVICIO DE REPARACION", "RECEPCION", vSQL, vOrden)
         Case "T"
           vSQL = "{POS_REPARACION_DETALLE.BOLETA_TRASLADO} = '" & vBoleta & "'"
           Call sbPosReportesSR("BoletaTraslado", "SERVICIO DE REPARACION", "TRASLADO A TALLER", vSQL, vOrden)
         Case "R"
             vSQL = "{POS_REPARACION_DETALLE.BOLETA_INGRESO} = '" & vBoleta & "'"
             Call sbPosReportesSR("BoletaRecibo", "SERVICIO DE REPARACION", "RECIBO DE TALLER", vSQL, vOrden)
         Case "E"
             vSQL = "{POS_REPARACION_DETALLE.BOLETA_ENTREGA} = '" & vBoleta & "'"
           Call sbPosReportesSR("BoletaEntrega", "SERVICIO DE REPARACION", "ENTREGA A CLIENTE", vSQL, vOrden)
       End Select
   
   Case "inventario"
       Select Case txtMovimiento.Tag
          Case "E"
              vSubTitulo = "SALIDAS"
          Case "S"
              vSubTitulo = "SALIDAS"
        End Select
        vSQL = "{PV_INVTRANSAC.BOLETA} = '" & vBoleta & "' AND {PV_INVTRANSAC.TIPO} = '" & txtMovimiento.Tag & "'"
        Call sbInvReportes("TransaccionBoleta", "BOLETA DE " & vSubTitulo, "", vSQL, vOrden)

End Select
End Sub
