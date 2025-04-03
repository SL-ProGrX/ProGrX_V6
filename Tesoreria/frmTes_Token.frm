VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_Token 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10155
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4932
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   8700
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
      FlatScrollBar   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.CheckBox chkPendientes 
      Height          =   372
      Left            =   6840
      TabIndex        =   5
      Top             =   840
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6583
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Mostrar solo solicitudes pendientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin MSComctlLib.Toolbar tlbCarga 
      Height          =   312
      Left            =   3840
      TabIndex        =   0
      Top             =   840
      Width           =   2628
      _ExtentX        =   4630
      _ExtentY        =   556
      ButtonWidth     =   1931
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cerrar"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "cargar datos"
            ImageKey        =   "IMG3"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reportes"
            Key             =   "Reportes"
            ImageKey        =   "IMG4"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Resumen"
                  Text            =   "Resumen"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Detallado"
                  Text            =   "Detallado"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTes_Token.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTes_Token.frx":169C2
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTes_Token.frx":2D384
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTes_Token.frx":424F6
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTes_Token.frx":57668
            Key             =   "IMG5"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtToken 
      Height          =   372
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Token"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tokens para trámites en Bancos"
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
      Height          =   372
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   6372
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmTES_Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
 vModulo = 9
 End Sub

Private Sub Form_Load()
 vModulo = 9
 Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Token Id", 2100
    .Add , , "Fecha", 2200
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Estado", 2100, vbCenter
    .Add , , "Pendiente", 2100, vbRightJustify
    .Add , , "Total", 2100, vbRightJustify
End With
 
 
Call sbLlenaLista
End Sub


Private Sub sbLlenaLista()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

strSQL = "select Top 100  Tok.ID_TOKEN, Tok.ESTADO , Tok.REGISTRO_FECHA , Tok.REGISTRO_USUARIO" _
       & "    , isnull( count(*) , 0) as 'Pendiente'  , isnull( sum(Tra.Monto), 0) as 'Monto'" _
       & " from TES_TOKENS Tok left join TES_TRANSACCIONES Tra on Tok.ID_TOKEN = Tra.ID_TOKEN and Tra.ESTADO = 'P'" _
       & " group by Tok.ID_TOKEN, Tok.ESTADO , Tok.REGISTRO_FECHA , Tok.REGISTRO_USUARIO" _
       & " order by Tok.registro_fecha desc"
Call OpenRecordSet(rs, strSQL)



Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , Trim(rs!id_token))
       itmX.SubItems(1) = Format(rs!REGISTRO_FECHA, "dd/mm/yyyy")
       itmX.SubItems(2) = UCase(rs!REGISTRO_USUARIO)
       
       If rs!Estado = "A" Then
          itmX.SubItems(3) = "Abierto"
       Else
          itmX.SubItems(3) = "Cerrado"
       End If
          
    If rs!Monto > 0 Then
        itmX.SubItems(4) = rs!Pendiente
    Else
        itmX.SubItems(4) = rs!Pendiente - 1
    End If
    itmX.SubItems(5) = Format(rs!Monto, "Standard")


rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtToken.Text = Item.Text

End Sub


Private Sub tlbCarga_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim strtToken As String

On Error GoTo vError

Select Case Button.Key
 Case "Cerrar"
    If Trim(txtToken.Text) <> "" Then
        strSQL = "select id_token from tes_tokens where estado = 'A' and id_token ='" & Trim(txtToken.Text) & "'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            strSQL = "update tes_tokens set estado = 'C' where id_token ='" & Trim(txtToken.Text) & "' "
            Call ConectionExecute(strSQL)
            MsgBox "Token cerrado satisfactoriamente...", vbInformation
        Else
            MsgBox "Este token ya esta cerrado", vbInformation
        End If
        rs.Close
       
        Call sbLlenaLista
    Else
       MsgBox "Debe digitar un token...", vbInformation
    End If

End Select

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Sub tlbCarga_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
  Case "Resumen"
    Call sbImprimeReportes("R")
  Case "Detallado"
    Call sbImprimeReportes("D")
End Select
End Sub

Private Sub sbImprimeReportes(pTipo As String)
Dim vFiltro As String


With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Banking"
    
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fxUsuario = '" & glogon.Usuario & "'"
    .Formulas(3) = "Token = '" & Trim(txtToken.Text) & "'"
          
    .Connect = glogon.ConectRPT
    vFiltro = ""
    .Formulas(5) = "SubTitulo = 'Todas las solicitudes'"
    
    If pTipo = "R" Then
        .Formulas(4) = "Titulo = 'Reporte Resumen de envio a Tesoreria'"
        .ReportFileName = SIFGlobal.fxPathReportes("Banking_ListadoTokenAgrupado.rpt")
    Else
       .Formulas(4) = "Titulo = 'Reporte Detallado de envio a Tesoreria'"
       .ReportFileName = SIFGlobal.fxPathReportes("Banking_ListadoTokenDetallado.rpt")
    End If
    vFiltro = "{Tes_transacciones.id_token} = '" & Trim(txtToken.Text) & "'"
    
    If chkPendientes.Value = vbChecked Then
        vFiltro = vFiltro & " and  {Tes_transacciones.Estado} = 'P'"
        .Formulas(5) = "SubTitulo = '--Solo Casos Pendientes de Desembolsar--'"
    End If
    .SelectionFormula = vFiltro
    .PrintReport
End With
 
End Sub

