VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmUS_Roles 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Roles de Seguridad"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   12870
   HelpContextID   =   1002
   Icon            =   "frmUS_Roles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbVincular 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   8400
      Width           =   12615
      _Version        =   1441793
      _ExtentX        =   22251
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnVincular 
         Height          =   375
         Index           =   0
         Left            =   9960
         TabIndex        =   9
         Top             =   360
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
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
         Picture         =   "frmUS_Roles.frx":08CA
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
         _ExtentY        =   661
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
         Text            =   "--"
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnVincular 
         Height          =   375
         Index           =   1
         Left            =   10440
         TabIndex        =   10
         Top             =   360
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
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
         Picture         =   "frmUS_Roles.frx":0FF1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vincluar Rol:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblRol 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   240
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6960
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
            Picture         =   "frmUS_Roles.frx":1595
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Roles.frx":168E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   12615
      _Version        =   524288
      _ExtentX        =   22251
      _ExtentY        =   11668
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
      MaxCols         =   492
      ScrollBars      =   2
      SpreadDesigner  =   "frmUS_Roles.frx":17C5
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1320
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Roles del Sistema"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmUS_Roles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnVincular_Click(Index As Integer)

If txtCliente.Tag = "" Then Exit Sub

Select Case Index
  Case 0 'Vincular
    strSQL = "update us_roles set cod_empresa  = " & txtCliente.Tag & " where cod_rol = '" & lblRol.Caption & "'"
  
  Case 1 'Desvincular
    strSQL = "update us_roles set cod_empresa  = null where cod_rol = '" & lblRol.Caption & "'"
End Select

Call ConectionExecute(strSQL)
If Index = 0 Then
        MsgBox "Cliente Vinculado al Rol, satisfactoriamente!", vbInformation
Else
        MsgBox "Cliente Desvinculado al Rol, satisfactoriamente!", vbInformation
End If

Call sbInicial

End Sub

Private Sub Form_activate()
vModulo = 13
End Sub

Private Sub sbInicial()

txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)


strSQL = "select R.cod_Rol,R.descripcion,R.activo" _
      & ", convert(varchar(10), isnull(R.cod_Empresa,0)) + '- ' + rtrim( isnull(C.Nombre_Largo,'General') ) as 'Cliente' " _
      & ", R.registro_Fecha,R.registro_Usuario,0" _
      & "  from US_Roles R left join PGX_Clientes C on R.cod_empresa = C.cod_Empresa" _
      & " Where R.descripcion like '%" & txtFiltro.Text & "%'" _
      & " order by R.descripcion"
vPaso = True
    Call sbCargaGrid(vGrid, 7, strSQL)
vPaso = False
End Sub


Private Sub Form_Load()

vModulo = 13

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from US_Roles " _
       & " where cod_Rol = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into US_Roles(cod_Rol,descripcion,activo,registro_fecha,registro_usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",Getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 4
  vGrid.Text = "0-General"
  vGrid.Col = 5
  vGrid.Text = fxFechaServidor
  vGrid.Col = 6
  vGrid.Text = glogon.Usuario
  
'  Call Bitacora("Registra", "Rol de Usuario: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update US_Roles set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where cod_Rol = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
' Call Bitacora("Modifica", "Rol de Usuario: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub



Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  Call sbCliente_Consulta
  txtCliente.Tag = gBusquedas.Resultado
  txtCliente.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sbInicial
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

vGrid.Row = Row

If Col = 7 Then
  gEntidad.Tipo = "R"
  
  vGrid.Col = 1
  gEntidad.Rol_Id = vGrid.Text
  vGrid.Col = 2
  gEntidad.Rol_Name = vGrid.Text
  
  frmUS_DerechosNew.Show vbModal

End If
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 3) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete US_Roles where cod_Rol = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Rol de Usuario: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If vPaso Then Exit Sub

vGrid.Row = NewRow
vGrid.Col = 5
If vGrid.Text <> "" Then
   vGrid.Col = 1
   lblRol.Caption = vGrid.Text
   vGrid.Col = 2
   lblRol.ToolTipText = vGrid.Text
   vGrid.Col = 4
   txtCliente.Text = vGrid.Text
   txtCliente.Tag = SIFGlobal.fxCodText(vGrid.Text)
   
   
End If

End Sub
