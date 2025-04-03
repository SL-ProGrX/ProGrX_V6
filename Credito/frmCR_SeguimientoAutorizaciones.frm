VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmCR_SeguimientoAutorizaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorizaciones de Solicitudes"
   ClientHeight    =   6624
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   8712
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6624
   ScaleWidth      =   8712
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAutorizar 
      Height          =   492
      Left            =   5520
      TabIndex        =   0
      Top             =   5880
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Autorizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_SeguimientoAutorizaciones.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   492
      Left            =   6840
      TabIndex        =   1
      Top             =   5880
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Reporte"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_SeguimientoAutorizaciones.frx":05DF
   End
   Begin TabDlg.SSTab ssTabMain 
      Height          =   4572
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8448
      _ExtentX        =   14901
      _ExtentY        =   8065
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Autorizaciones"
      TabPicture(0)   =   "frmCR_SeguimientoAutorizaciones.frx":0D9B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtNotas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtOperacion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDetalle"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtDetalle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2265
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   900
         Width           =   6855
      End
      Begin VB.TextBox txtOperacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   384
         Left            =   1440
         TabIndex        =   4
         Top             =   540
         Width           =   1812
      End
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1185
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3300
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   3300
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorización de la Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   2160
      TabIndex        =   8
      Top             =   300
      Width           =   5772
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCR_SeguimientoAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As Boolean
Dim vPorcentaje As Double, vMontoRefunde As Currency, vAhorros As Currency
Dim vMontoCredito As Currency, vCedula As String

'Verifica que si el credito es sobre ahorros, la autorizacion no sobre pase el 100% de los ahorros

vResultado = True

strSQL = "select garantia,montoapr,cedula from reg_creditos where id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

If rs!Garantia = "A" Then
  vPorcentaje = 1 '100% de los ahorros
   
   vMontoCredito = rs!montoapr
   vCedula = rs!Cedula
  
   rs.Close
    
   'Saldos Sobre Ahorros de creditos vigentes
   strSQL = "select  isnull(sum(saldo),0) as Monto" _
          & " from reg_creditos where estado = 'A' and saldo > 0 and garantia = 'A'" _
          & " and cedula = '" & vCedula & "'"
  
   Call OpenRecordSet(rs, strSQL)
     vMontoRefunde = rs!Monto
   rs.Close
  
  
    ' Total de Saldos Sobre Ahorros menos las refundiciones sobre ahorros
    ' Queda el pendiente en sobre ahorros
   strSQL = "select isnull(sum(R.monto),0) as Refunde" _
           & " from refundiciones R inner join reg_creditos C on R.id_solicitud = C.id_solicitud" _
           & " where R.id_solicitudr = " & txtOperacion.Text & " and C.garantia = 'A'"
   Call OpenRecordSet(rs, strSQL)
     vMontoRefunde = vMontoRefunde - rs!Refunde
   rs.Close
    
    ' Ahorros Actuales
    strSQL = " select (isnull(AHORRO,0) + isnull(Capitaliza,0)) as Ahorro" _
           & " from ahorro_consolidado where cedula = '" & vCedula & "'"
    Call OpenRecordSet(rs, strSQL)
      vAhorros = rs!ahorro
   
    If (vMontoCredito + vMontoRefunde) > (vAhorros * vPorcentaje) Then
       vResultado = False
    End If
    
  
End If
rs.Close

fxValida = vResultado


End Function


Private Sub cmdAutorizar_Click()
Dim strSQL As String

On Error GoTo vError

If txtDetalle.Tag = "N" Then
    MsgBox "Es solicitud no ha pasado la validación, verifique...", vbExclamation
    Exit Sub
End If


If Len(Trim(txtNotas.Text)) = 0 Then
    MsgBox "No se ha especificado ninguna nota para la autorización, verifique...", vbExclamation
    Exit Sub
End If


'If Not fxValida Then
'   MsgBox "La autorización no puede sobrepasar el 100% de los ahorros de la persona como garantía del crédito...!", vbExclamation
'   Exit Sub
'End If

strSQL = "update reg_creditos set autoriza_user = '" & glogon.Usuario & "',autoriza_fecha = dbo.MyGetdate()" _
       & ",Autoriza_nota = '" & txtNotas.Text _
       & "' where id_solicitud = " & txtOperacion
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Autorizacion Solicitud : " & txtOperacion.Text)
'Tags de Seguimiento
Call sbCrdOperacionTags(txtOperacion.Text, txtOperacion.Tag, "S09", "", txtNotas.Text)

MsgBox "Solicitud Autorizada Satisfactoriamente...", vbInformation

Call txtOperacion_Change

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3


Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

Call txtOperacion_Change

Call Formularios(Me)
Call RefrescaTags(Me)

ssTabMain.Tab = 0
Call ssTabMain_Click(0)

cmdAutorizar.Enabled = cmdAutorizar.Enabled


End Sub


Private Sub ssTabMain_Click(PreviousTab As Integer)

cmdAutorizar.Visible = False
cmdReporte.Visible = False

If ssTabMain.Tab = 0 Then
  cmdAutorizar.Visible = True
Else
  cmdReporte.Visible = True
End If

End Sub


Private Sub txtOperacion_Change()
    txtDetalle.Text = ""
    txtDetalle.Tag = "N"
    txtNotas.Text = ""

End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsulta

Exit Sub

vError:

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

txtDetalle.Text = ""
txtDetalle.Tag = "N"
txtNotas.Text = ""


strSQL = "Select R.id_solicitud,R.codigo,R.cedula,S.nombre,R.montoapr,R.plazo,R.int, R.cuota,R.cod_destino" _
       & ",D.descripcion as DestinoX,C.descripcion as LineaX,R.userrec,R.fechasol,R.observacion,R.garantia" _
       & " from reg_creditos R inner join socios S on R.cedula = S.cedula" _
       & " inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " inner join catalogo_Destinos D on R.cod_destino = D.cod_Destino" _
       & " where R.autoriza_user is null and R.estadosol = 'R' and R.id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then

  txtDetalle = txtDetalle & vbCrLf & "# Solicitud : " & vbTab & rs!id_solicitud
  txtDetalle = txtDetalle & vbCrLf & "Línea       : " & vbTab & rs!Codigo & " - " & rs!LineaX
  txtDetalle = txtDetalle & vbCrLf & "Destino     : " & vbTab & rs!cod_destino & " - " & rs!DestinoX
  txtDetalle = txtDetalle & vbCrLf & "Cédula      : " & vbTab & Trim(rs!Cedula) & " - " & rs!Nombre & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & "Monto       : " & vbTab & Format(rs!montoapr, "Standard")
  txtDetalle = txtDetalle & vbCrLf & "Plazo       : " & vbTab & rs!Plazo
  txtDetalle = txtDetalle & vbCrLf & "Tasa        : " & vbTab & rs!Int
  txtDetalle = txtDetalle & vbCrLf & "Cuota       : " & vbTab & Format(rs!Cuota, "Standard") & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & "Garantía    : " & vbTab & fxGarantia(rs!Garantia) & vbCrLf

  txtDetalle = txtDetalle & vbCrLf & "Fecha       : " & vbTab & Format(rs!FechaSol, "dd/mm/yyyy")
  txtDetalle = txtDetalle & vbCrLf & "Usuario     : " & vbTab & rs!userRec & vbCrLf

  txtDetalle = txtDetalle & vbCrLf & "Notas : " & rs!observacion & ""

  txtDetalle.Tag = "S"
  txtOperacion.Tag = rs!Codigo
  
End If
rs.Close

Me.MousePointer = vbDefault

If txtDetalle.Tag = "N" Then
   txtDetalle.ForeColor = vbRed
   MsgBox " La Solicitud no cumple con alguno(s) de los siguientes parámetros:" _
          & vbCrLf & " 1. No se encuentra recibida" & vbCrLf & " 2. No Existe la Solicitud (Credito)" _
          & vbCrLf & " 3. Ya se encuentra Autorizada?", vbExclamation
Else
   txtDetalle.ForeColor = vbBlue
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 Call txtOperacion_Change

End Sub

