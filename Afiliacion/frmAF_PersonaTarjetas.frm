VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_PersonaTarjetas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Tarjetas"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2892
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   7932
      _Version        =   1441793
      _ExtentX        =   13991
      _ExtentY        =   5101
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   612
      Index           =   0
      Left            =   6480
      TabIndex        =   12
      Top             =   1920
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_PersonaTarjetas.frx":0000
   End
   Begin XtremeSuiteControls.CheckBox chkValidaTarjeta 
      Height          =   372
      Left            =   6480
      TabIndex        =   11
      Top             =   1440
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Valida Tarjeta?"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   6000
      Top             =   1560
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   8280
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":0720
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":0E17
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgCardTypeMini 
      Left            =   7680
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":17D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":462A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":7563
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":A59F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgCardTypes 
      Left            =   7080
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   176
      ImageHeight     =   107
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":D3FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":10251
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":1318A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":161C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_PersonaTarjetas.frx":19022
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtTarjetaNumero 
      Height          =   312
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6583
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtVenceMes 
      Height          =   312
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   372
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.FlatEdit txtVenceAnio 
      Height          =   312
      Left            =   2760
      TabIndex        =   9
      Top             =   1800
      Width           =   372
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   612
      Index           =   1
      Left            =   7080
      TabIndex        =   13
      Top             =   1920
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_PersonaTarjetas.frx":1BE29
   End
   Begin XtremeSuiteControls.FlatEdit txtSecurityCode 
      Height          =   312
      Left            =   2160
      TabIndex        =   10
      Top             =   2160
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Alignment       =   2
      Locked          =   -1  'True
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(Opcional)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3240
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Code:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image imgCardType 
      Height          =   645
      Left            =   4560
      Picture         =   "frmAF_PersonaTarjetas.frx":1C3CD
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
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
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MM / AA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Vencimiento:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Tarjeta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjetas Registradas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1880
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmAF_PersonaTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnTool_Click(Index As Integer)
Dim strSQL As String, vVence As Date, vTipoMov As String

On Error GoTo vError

'(@ClienteCod int, @Cedula varchar(30), @Tarjeta varchar(20), @Vence datetime
'                    ,@Code varchar(10), @TipoMov char(1), @Usuario varchar(30) , @Token varchar(30) = '')


strSQL = ""


If Index = 0 Then
    'Guardar
    Select Case False
        Case IsNumeric(txtTarjetaNumero.Text)
            strSQL = "La Tarjeta no es válida!"
        Case IsNumeric(txtVenceAnio.Text)
            strSQL = "Año de Vencimiento no es válido!"
        Case IsNumeric(txtVenceMes.Text)
            strSQL = "Mes de Vencimiento no es válido!"
    End Select
    
    If Len(strSQL) > 0 Then
       MsgBox strSQL, vbExclamation
       Exit Sub
    End If
    
    If Not fxTarjetaValida(txtTarjetaNumero.Text) And chkValidaTarjeta.Value = vbChecked Then
       MsgBox "Tarjeta no es Válida, verfique!", vbExclamation
       Exit Sub
    End If

End If

vVence = CDate("20" + txtVenceAnio.Text & "/" & Format(txtVenceMes.Text, "00") & "/01")

Select Case Index
   Case 0 'Guardar
        vTipoMov = "A"
        Call Bitacora("Registra", "Tarjeta: " & txtTarjetaNumero.Text & " Id:" & GLOBALES.gCedulaActual)
     
   Case 1 'Borrar
        vTipoMov = "E"
        Call Bitacora("Elimina", "Tarjeta: " & txtTarjetaNumero.Text & " Id:" & GLOBALES.gCedulaActual)
End Select
     
strSQL = "exec  spAFI_PersonaTarjetas_Registro " & gPortal.Empresa_Id & ",'" & GLOBALES.gCedulaActual & "','" _
       & txtTarjetaNumero.Text & "','" & Format(vVence, "yyyy/mm/dd") & "','" & SIFGlobal.fxStringCifrado(txtSecurityCode.Text) & "','" & vTipoMov & "','" _
       & glogon.Usuario & "',''"
Call ConectionExecute(strSQL)

Call sbTarjetasLlenaList

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Me.Caption = "[Identificación : " & GLOBALES.gCedulaActual & "]"

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
  .Clear
  .Add , , "Número", 3500
  .Add , , "Vence", 1500, vbCenter
  .Add , , "Code", 1500, vbCenter
  .Add , , "Tipo", 1500, vbCenter
End With

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxTarjetaMask(pTarjeta As String) As String
Dim vResultado As String

vResultado = Mid(pTarjeta, 1, 4) & "********" & Right(pTarjeta, 4)

fxTarjetaMask = vResultado

End Function

Private Sub sbTarjetasLlenaList()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pImagen As Integer

On Error GoTo vError

txtTarjetaNumero.Text = ""
txtTarjetaNumero.SetFocus
txtSecurityCode.Text = ""
txtVenceAnio.Text = ""
txtVenceMes.Text = ""
Set imgCardType.Picture = Nothing

strSQL = "exec spAFI_PersonaTarjetas_Consulta " & gPortal.Empresa_Id & ",'" & GLOBALES.gCedulaActual & "',''"
Call OpenRecordSet(rs, strSQL)

With lsw.ListItems
   .Clear
   Do While Not rs.EOF
    
        Select Case UCase(rs!Tarjeta_Tipo)
            Case "VISA"
               pImagen = 1
            Case "MASTERCARD"
               pImagen = 2
            Case "AMERICAN EXPRESS"
               pImagen = 3
            Case "DISCOVER"
               pImagen = 4
        End Select
    
    Set itmX = .Add(, , rs!Tarjeta_Mask)
        itmX.SubItems(1) = Format(rs!Tarjeta_Vence, "MM/YY")
        itmX.SubItems(2) = "****"
        itmX.SubItems(3) = rs!Tarjeta_Tipo
        
        
        itmX.ListSubItems(1).Tag = rs!tarjeta_Numero
        itmX.ListSubItems(2).Tag = rs!Tarjeta_Vence
        itmX.ListSubItems(3).Tag = pImagen
        
        itmX.ToolTipText = rs!Tarjeta_Code
    rs.MoveNext
   Loop
   rs.Close
End With


Exit Sub

vError:
  lsw.ListItems.Clear
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

With Item

  txtTarjetaNumero.Text = .ListSubItems(1).Tag
  txtVenceAnio.Text = Right(Year(.ListSubItems(2).Tag), 2)
  txtVenceMes.Text = Format(Month(.ListSubItems(2).Tag), "00")
  
  
  txtSecurityCode.Text = .ToolTipText

  Set imgCardType.Picture = imgCardTypes.ListImages.Item(.ListSubItems(3).Tag).Picture

End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbTarjetasLlenaList
End Sub


Private Sub txtSecurityCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtTarjetaNumero.SetFocus

End Sub

Private Sub txtTarjetaNumero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtVenceMes.SetFocus
End Sub

Private Sub txtTarjetaNumero_KeyUp(KeyCode As Integer, Shift As Integer)
Dim vTipo As String

vTipo = fxTarjetaTipo(txtTarjetaNumero.Text)
Select Case UCase(vTipo)
   Case "VISA"
      Set imgCardType.Picture = imgCardTypes.ListImages.Item(1).Picture
   Case "MASTERCARD"
      Set imgCardType.Picture = imgCardTypes.ListImages.Item(2).Picture
   Case "AMERICAN EXPRESS"
      Set imgCardType.Picture = imgCardTypes.ListImages.Item(3).Picture
   Case "DISCOVER"
      Set imgCardType.Picture = imgCardTypes.ListImages.Item(4).Picture
   Case Else
       Set imgCardType.Picture = Nothing
End Select
End Sub


Private Sub txtVenceAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtSecurityCode.SetFocus

End Sub


Private Sub txtVenceMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtVenceAnio.SetFocus

End Sub
