VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCajas_ReporteCierres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Cierre Caja"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11145
   Icon            =   "frmCajas_ReporteCierres.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11145
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5172
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   10932
      _Version        =   1441792
      _ExtentX        =   19283
      _ExtentY        =   9123
      _StockProps     =   77
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkModoSupervisor 
      Height          =   372
      Left            =   4800
      TabIndex        =   10
      Top             =   1560
      Width           =   3372
      _Version        =   1441792
      _ExtentX        =   5948
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Aperturas/Cierres (Saldos Abiertos)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnCierre 
      Height          =   612
      Left            =   9000
      TabIndex        =   3
      Top             =   1200
      Width           =   1692
      _Version        =   1441792
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Forzar Cierre"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   16
      Picture         =   "frmCajas_ReporteCierres.frx":0ECA
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_ReporteCierres.frx":18B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_ReporteCierres.frx":19ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCaja 
      Height          =   312
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtApertura 
      Height          =   312
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCajaDesc 
      Height          =   312
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   5532
      _Version        =   1441792
      _ExtentX        =   9758
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   492
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin MSComctlLib.Toolbar tblAplicar 
         Height          =   312
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   556
         ButtonWidth     =   1799
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Reportes del Cierre"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Resumen"
                     Text            =   "Resumen"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Cierre"
                     Text            =   "Informe de Cierre"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Movimientos"
                     Text            =   "Movimientos"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Cierres de Cajas"
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
      Height          =   372
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   6612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Apertura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   11172
   End
End
Attribute VB_Name = "frmCajas_ReporteCierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCierreCiego As Boolean, vPaso As Boolean

Private Sub btnCierre_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vAplicado As Boolean

If Not IsNumeric(txtApertura.Text) Then Exit Sub

On Error GoTo vError

vAplicado = False

strSQL = "select count(*) as 'Existe'" _
       & " From CAJAS_APERTURAS_MAIN" _
       & " where  COD_CAJA = '" & txtCaja.Text & "' and cod_apertura = " & txtApertura.Text _
       & " and Estado = 'A'"
       
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 1 Then
  
  strSQL = "exec spCajas_Cierre_Forzado '" & txtCaja.Text & "'," & txtApertura.Text & ",'" & glogon.Usuario & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Aplica", "Cierre Forzado a Caja: " & txtCaja.Text & " AP: " & txtApertura.Text)
    
  vAplicado = True
End If
rs.Close

Me.MousePointer = vbDefault

If vAplicado Then
    MsgBox "Cierre de Caja: " & txtCaja.Text & " Apertura: " & txtApertura.Text & ", Realizado Satisfactoriamente!", vbInformation
    Call txtCaja_LostFocus
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Sub Form_Load()
 vModulo = 5
 
 Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "No.Apertura", 1500
    .Add , , "Fecha", 2100
    .Add , , "Usuario", 1850, vbCenter
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Cierre [Fecha]", 2100
    .Add , , "Cierre [Usuario]", 1850, vbCenter
  End With
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count = 0 Or vPaso Then Exit Sub

txtApertura.Text = Item.Text
txtApertura.Tag = Mid(Item.SubItems(3), 1, 1)
End Sub



Private Sub tblAplicar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String

'Modo de Supervision
If chkModoSupervisor.Value = vbChecked Then
   vCierreCiego = False
End If

'Aplica un Cierre Preliminar de los datos para ver el informe
If txtApertura.Tag = "Abierta" Then
   strSQL = "exec spCajas_CierreCajaMain '" & txtCaja.Text & "'," & txtApertura.Text _
       & ",'" & glogon.Usuario & "',1"
   Call ConectionExecute(strSQL)
End If

Select Case ButtonMenu.Key
  Case "Resumen"
     Call sbCajasCierreReportes(txtCaja.Text, txtApertura.Text, "Resumen", vCierreCiego)
  Case "Cierre"
     Call sbCajasCierreReportes(txtCaja.Text, txtApertura.Text, "Cierre", vCierreCiego)
  Case "Movimientos"
     Call sbCajasCierreReportes(txtCaja.Text, txtApertura.Text, "Movimientos", vCierreCiego)
End Select

End Sub

Private Sub sbAperturaConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Me.MousePointer = vbHourglass

On Error GoTo vError

txtApertura.Tag = "X"

strSQL = "select *" _
       & " From CAJAS_APERTURAS_MAIN" _
       & " where  COD_CAJA = '" & txtCaja.Text & "' and cod_apertura = " & txtApertura.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtApertura.Text = rs!cod_Apertura
    txtApertura.Tag = rs!Estado
Else
    MsgBox "La apertura consultada (No." & txtApertura.Text & ")no existe verifique!", vbExclamation
    txtApertura.Text = 0
    txtApertura.Tag = "X"
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub txtApertura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtCajaDesc.SetFocus
End Sub

Private Sub txtApertura_LostFocus()
 Call sbAperturaConsulta
End Sub

Private Sub txtCaja_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtCajaDesc.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado2 = ""
    gBusquedas.Resultado = ""
    txtCaja = ""
    txtCajaDesc = ""
    gBusquedas.Columna = "cod_caja"
    gBusquedas.Orden = "cod_caja"
    gBusquedas.Consulta = "Select cod_caja,descripcion From cajas_definicion"
    
    frmBusquedas.Show vbModal
    
    txtCaja = Trim(gBusquedas.Resultado)
    txtCajaDesc = gBusquedas.Resultado2
        
    If gBusquedas.Resultado <> "" Then txtCajaDesc.SetFocus
    
End If


End Sub

Private Sub txtCaja_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear
txtApertura.Text = 0
txtApertura.Tag = "X"
vCierreCiego = True

strSQL = "select descripcion,cierre_tipo from cajas_Definicion where cod_Caja = '" & txtCaja.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  If rs!Cierre_Tipo = "A" Then vCierreCiego = False
  txtCajaDesc.Text = rs!Descripcion
End If
rs.Close


strSQL = "select Top 30 *" _
       & " From CAJAS_APERTURAS_MAIN" _
       & " where  COD_CAJA = '" & txtCaja.Text & "'" _
       & " order by COD_APERTURA desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_Apertura)
      itmX.SubItems(1) = rs!Apertura_Fecha
      itmX.SubItems(2) = rs!Apertura_Usuario
      itmX.SubItems(3) = IIf((rs!Estado = "C"), "Cerrada", "Abierta")
      itmX.SubItems(4) = rs!Cierre_Fecha & ""
      itmX.SubItems(5) = rs!Cierre_Usuario & ""
      
  If txtApertura.Text = 0 Then
        txtApertura.Text = rs!cod_Apertura
        txtApertura.Tag = rs!Estado
  End If
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub txtCajaDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado2 = ""
    gBusquedas.Resultado = ""
    txtCaja = ""
    txtCajaDesc = ""
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "Select cod_caja,descripcion From cajas_definicion"
    
    frmBusquedas.Show vbModal
    
    txtCaja = Trim(gBusquedas.Resultado)
    txtCajaDesc = gBusquedas.Resultado2
    
    If gBusquedas.Resultado <> "" Then
       Call txtCaja_LostFocus
    End If
    
End If

End Sub
