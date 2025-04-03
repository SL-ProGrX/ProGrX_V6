VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Apelacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fosol: Apelaciones"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab 
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Apelaciones"
      TabPicture(0)   =   "frmFSL_Apelacion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vgGrid"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Registro"
      TabPicture(1)   =   "frmFSL_Apelacion.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dtpFechaApelacion"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtObservaciones"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtNombrePresenta"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtNombrePresenta 
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   720
         Width           =   5565
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   2235
         Left            =   120
         MaxLength       =   1500
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2100
         Width           =   8805
      End
      Begin FPSpreadADO.fpSpread vgGrid 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   14
         Top             =   660
         Width           =   8655
         _Version        =   524288
         _ExtentX        =   15266
         _ExtentY        =   6165
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   5
         SpreadDesigner  =   "frmFSL_Apelacion.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpFechaApelacion 
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   139067393
         CurrentDate     =   41043
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Presenta Apelación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Apelación"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Notas sobre Apelación"
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1740
         Width           =   2415
      End
   End
   Begin VB.TextBox txtCedula 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   9
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   600
      Width           =   2070
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   4485
   End
   Begin VB.TextBox txtExpediente 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "Expediente"
      Top             =   960
      Width           =   2055
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   9390
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlbMantenimiento"
      MinHeight1      =   330
      Width1          =   4095
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbMenu 
         Height          =   330
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   13050
         _ExtentX        =   23019
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo Expediente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Key             =   "Guardar"
               Object.ToolTipText     =   "Guarda los datos"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               Object.ToolTipText     =   "imprimi Boleta del Expediente"
               ImageIndex      =   9
            EndProperty
         EndProperty
         Begin VB.Label Label5 
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   7
            Top             =   2160
            Width           =   5055
         End
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   5
         ToolTipText     =   "Exposición a Riesgo de la persona"
         Top             =   0
         Width           =   6660
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   30
         Width           =   5265
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   165
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   30
         Width           =   2145
      End
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   8640
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":071D
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":1588F
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":2AA01
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":3FB73
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":54CE5
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":555BF
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":6A731
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":7F8A3
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Apelacion.frx":8017D
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Asociado"
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
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Expediente"
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
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmFSL_Apelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExpediente As Long

Private Sub Form_Activate()
  vModulo = 22
End Sub

Private Sub Form_Load()
  vModulo = 22
  txtCedula.Text = GLOBALES.gTag
  txtNombre.Text = GLOBALES.gTag2
  txtExpediente.Text = GLOBALES.gTag3
  vExpediente = txtExpediente.Text
  ssTab.Tab = 0
  dtpFechaApelacion.Value = fxFechaServidor
  
  Call sbCargaApelaciones
End Sub

Private Sub sbLimpiar()
 txtNombrePresenta.Text = Empty
 txtObservaciones.Text = Empty
 dtpFechaApelacion.Value = fxFechaServidor
End Sub

Private Function fxExisteApelacion() As Boolean

strSQL = "Select count(COD_EXPEDIENTE) as Conteo from FSL_APELACIONES " _
       & " where COD_EXPEDIENTE = " & txtExpediente.Text & ""
rs.Open strSQL, glogon.Conection, adOpenStatic
 
If rs!Conteo <= 0 Then
  fxExisteApelacion = False
End If
vExpediente = txtExpediente.Text

rs.Close
  
End Function

Private Sub sbModificaExpediente()
  strSQL = "Update FSL_EXPEDIENTES set ESTADO = 'APL' where COD_EXPEDIENTE = " & vExpediente & ""
  glogon.Conection.Execute strSQL
End Sub

Private Sub sbGuardaApelacion()
   strSQL = "Insert FSL_APELACIONES (COD_EXPEDIENTE, FECHA_APELACION, NOMBRE_PRESENTA, OBSERVACIONES,REGISTRA_USUARIO,REGISTRA_FECHA)" _
          & " values (" & vExpediente & ",'" & Format(dtpFechaApelacion.Value, "yyyymmdd") & "','" & txtNombrePresenta.Text & "'" _
          & ", '" & txtObservaciones.Text & "','" & glogon.Usuario & "',getdate())"
   glogon.Conection.Execute strSQL
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
   Select Case ssTab.Tab
      Case 0
        Call sbCargaApelaciones
        dtpFechaApelacion = fxFechaServidor
   End Select
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Nuevo"
     Call sbLimpiar
     ssTab.Tab = 1
     
  Case "Guardar"
     If fxExisteApelacion = False Then
        Call sbGuardaApelacion
        Call sbModificaExpediente
        MsgBox "Apelación se registro satisfactoriamente!", vbExclamation
        
'       strSQL = "exec spSifRegistraTags '" & txtCedula.Text & "','" & txtExpediente.Text & "', " _
'              & " '','A04','" & glogon.Usuario & "','" & txtObservaciones.Text & "'," & txtExpediente.Text & ""
'
'       glogon.Conection.Execute strSQL
        
     End If
  Case "Imprimir"
       
       
End Select

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

Dim vCedTemp As String

On Error GoTo vError

'Busca primer en el Maestro de Socios, de lo contrario revisa si es una operacion
' y regresa la cedula de la operacion
vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 strSQL = "select coalesce(count(*),0) as Existe from socios where cedula = '" & txtCedula & "'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then
   rs.Close
   strSQL = "select cedula from reg_creditos where id_solicitud = " & txtCedula
   rs.Open strSQL, glogon.Conection, adOpenStatic
   If Not rs.EOF And Not rs.BOF Then
      vCedTemp = Trim(rs!Cedula)
   End If
 End If
 rs.Close

If vCedTemp = "" Then
  Call sbConsulta(txtCedula.Text)
Else
  Call sbConsulta(vCedTemp)
End If
End If

If KeyCode = vbKeyF4 Then Call sbBusqueda

Call sbCargaApelaciones

Exit Sub
vError:
    MsgBox Err.Description, vbCritical
End Sub

'Carga las apelaciones presentadas segun el expediente
Private Sub sbCargaApelaciones()
Dim i As Integer
On Error GoTo vError

strSQL = "Select A.CONSECUTIVO, A.FECHA_APELACION, A.OBSERVACIONES, A.REGISTRA_USUARIO,E.CEDULA,S.NOMBRE" _
       & " from FSL_APELACIONES A" _
       & " inner join FSL_EXPEDIENTES E on A.COD_EXPEDIENTE = E.COD_EXPEDIENTE " _
       & " inner join SOCIOS S on E.CEDULA = S.CEDULA " _
       & " Where A.COD_EXPEDIENTE = " & vExpediente & ""
rs.Open strSQL, glogon.Conection, adOpenStatic

With vgGrid
 .MaxRows = 1
 .Row = .MaxRows
 For i = 1 To .MaxCols
   .Col = i
   .Text = ""
 Next i

 
 Do While Not rs.EOF
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
 
  .Row = .MaxRows
       
  .Col = 1
  .Text = Format(rs!consecutivo, "standard")
        
  .Col = 2
  .Text = Format(rs!FECHA_APELACION, "dd/mm/yyyy")
      
  .Col = 3
  .Text = rs!Nombre
      
  .Col = 4
  .Text = rs!observaciones
      
  .Col = 5
  .Text = IIf(IsNull(rs!Registra_Usuario), "", rs!Registra_Usuario)
  
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  rs.MoveNext
 Loop
End With
rs.Close

Exit Sub

vError:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub sbConsulta(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset
Dim vEstadoSocio As String
     
vFianzas = False
    
strSQL = "select S.cedula as CedulaX,S.nombre" _
       & " from socios S " _
       & " Where S.cedula = '" & Trim(pCedula) & "'"

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
 
If Not rs.EOF And Not rs.BOF Then
   
   txtCedula.Text = Trim(rs!Cedulax & "")
   txtNombre.Text = rs!Nombre & ""
   rs.Close
       
   ssTab.Tab = 0
 
 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
  
End Sub

Private Sub sbBusqueda()
On Error GoTo vError
  gBusquedas.Convertir = "N"
  gBusquedas.Consulta = "Select cedula,nombre from SOCIOS"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  frmBusquedas.Show vbModal
  txtCedula = Trim(gBusquedas.Resultado)
  gBusquedas.Consulta = ""
  gBusquedas.Columna = ""
  gBusquedas.Orden = ""
  gBusquedas.Resultado = ""
  If Trim(txtCedula) <> "" Then
       Call sbConsulta(txtCedula)
  End If

Exit Sub
vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub txtExpediente_KeyPress(KeyAscii As Integer)
  If (IsNumeric(Chr(KeyAscii)) <> True) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
     KeyAscii = 0
  End If
  
  If KeyAscii = 13 Then
    vExpediente = txtExpediente.Text
    Call sbCargaApelaciones
  End If
End Sub

Private Sub txtNombrePresenta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservaciones.SetFocus
End Sub
