VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "Codejock.Controls.v19.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "Codejock.ShortcutBar.v19.2.0.ocx"
Begin VB.Form frmCR_PlazosBonificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plazos de Bonficación por Membresía"
   ClientHeight    =   8460
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   13104
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   13104
   Begin XtremeSuiteControls.FlatEdit txtPlan 
      Height          =   492
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   492
      _Version        =   1245186
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7092
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   13092
      _Version        =   1245186
      _ExtentX        =   23093
      _ExtentY        =   12509
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "Definición"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "Label3(0)"
      Item(0).Control(1)=   "txtDescripcion"
      Item(0).Control(2)=   "Label3(1)"
      Item(0).Control(3)=   "txtNotas"
      Item(0).Control(4)=   "chkActivo"
      Item(0).Control(5)=   "tlb"
      Item(0).Control(6)=   "gbMain"
      Item(1).Caption =   "Bonificación"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Item(2).Caption =   "Asignación"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lsw"
      Item(2).Control(1)=   "lbl"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6492
         Left            =   3360
         TabIndex        =   5
         Top             =   720
         Width           =   6372
         _Version        =   1245186
         _ExtentX        =   11239
         _ExtentY        =   11451
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
         Checkboxes      =   -1  'True
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   3972
         Left            =   -69760
         TabIndex        =   6
         Top             =   3000
         Visible         =   0   'False
         Width           =   12612
         _Version        =   1245186
         _ExtentX        =   22246
         _ExtentY        =   7006
         _StockProps     =   79
         Caption         =   "Planes Registrados: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswPlanes 
            Height          =   3492
            Left            =   0
            TabIndex        =   7
            Top             =   360
            Width           =   12612
            _Version        =   1245186
            _ExtentX        =   22246
            _ExtentY        =   6159
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
         End
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   372
         Left            =   -58960
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   732
         _Version        =   1245186
         _ExtentX        =   1291
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Activo?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   -66400
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1245186
         _ExtentX        =   12721
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6492
         Left            =   -68800
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   11052
         _Version        =   524288
         _ExtentX        =   19495
         _ExtentY        =   11451
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
         SpreadDesigner  =   "frmCR_PlazosBonificacion.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   -66400
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   3828
         _ExtentX        =   6752
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
               Key             =   "Reportes"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1632
         Left            =   -66400
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1245186
         _ExtentX        =   12721
         _ExtentY        =   2879
         _StockProps     =   77
         ForeColor       =   0
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   0
         Left            =   -67960
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1245186
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripción"
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   1
         Left            =   -67960
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1245186
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption lbl 
         Height          =   312
         Left            =   3360
         TabIndex        =   13
         Top             =   360
         Width           =   6372
         _Version        =   1245186
         _ExtentX        =   11239
         _ExtentY        =   550
         _StockProps     =   14
         Caption         =   "Garantías asiganadas?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   492
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Plan de Bonificación"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmCR_PlazosBonificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vConsultaActiva As Integer, vNode As Node
Dim vEditar As Boolean, vScroll As Boolean, vPaso As Boolean

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_Plazo_Bono from CRD_PLAZO_BONO"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_Plazo_Bono > '" & txtPlan.Text & "' order by cod_Plazo_Bono asc"
    Else
       strSQL = strSQL & " where cod_Plazo_Bono < '" & txtPlan.Text & "' order by cod_Plazo_Bono desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtPlan.Text = rs!cod_Plazo_Bono
      Call sbConsulta(txtPlan.Text)
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

Private Sub Form_Load()

vModulo = 3

 vEditar = False

 
 Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

 tcMain.Item(0).Selected = True
 
 With lswPlanes.ColumnHeaders
    .Clear
    .Add , , "Plan", 1200
    .Add , , "Descripción", 3500
    .Add , , "Notas", 2500
    .Add , , "Activo?", 1100, vbCenter
    .Add , , "Usuario", 1600
    .Add , , "Registro", 2100
 End With
 
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Descripción", 6272
 End With
 
 lsw.Checkboxes = True
 lsw.HideColumnHeaders = True
 
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 
 lsw.Enabled = cmdActualiza.Enabled
 vGrid.Enabled = cmdActualiza.Enabled
 
 Call sbLimpia

End Sub


Private Sub sbLimpia(Optional pSoloLista As Boolean = False)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Select Case tcMain.SelectedItem
  Case 0 'Remesas
     If Not pSoloLista Then
             txtPlan.Text = ""
             
             txtDescripcion.Text = ""
             txtNotas.Text = ""
            
             chkActivo.Value = vbChecked
     End If
     
     strSQL = "select * from CRD_PLAZO_BONO order by cod_Plazo_Bono"
     lswPlanes.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswPlanes.ListItems
            Set itmX = .Add(, , rs!cod_Plazo_Bono)
                itmX.SubItems(1) = rs!Descripcion
                itmX.SubItems(2) = rs!Notas
                itmX.SubItems(3) = IIf((rs!Activo = 1), "Activo", "Inactivo")
                itmX.SubItems(4) = rs!registro_usuario & ""
                itmX.SubItems(5) = rs!registro_fecha & ""
       End With
       rs.MoveNext
     Loop
     rs.Close
     
  Case 1 'Tabla
   
  Case 2 'Asignacion
 End Select

End Sub


Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""
fxVerifica = True

If txtPlan.Text = "" Then vMensaje = vMensaje & " - Especifique un código del Plan de Bonificación" & vbCrLf
If txtDescripcion.Text = "" Then vMensaje = vMensaje & " - Especifique una descripción del Plan" & vbCrLf
If txtNotas.Text = "" Then vMensaje = vMensaje & " - Especifique una descripción del Plan" & vbCrLf


If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation
   fxVerifica = False
End If


End Function



Private Sub sbConsulta(pPlan As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error Resume Next

strSQL = "select * from CRD_PLAZO_BONO where cod_Plazo_Bono = '" & pPlan & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   vEditar = True
   
   Call sbToolBar(tlb, "activo")
   Call sbLimpia
   
   
   txtPlan.Text = rs!cod_Plazo_Bono
   txtDescripcion.Text = rs!Descripcion
   txtNotas.Text = rs!Notas
   chkActivo.Value = rs!Activo
   
   vCodigo = Trim(txtPlan)
    
  Else
   
   If vEditar = True Then
        vEditar = False
        Call sbToolBar(tlb, "nuevo")
        Call sbLimpia
        txtPlan.SetFocus
   End If

End If
rs.Close

End Sub

Private Sub sbBorrar()

End Sub


Private Sub sbGuardar()
Dim strSQL As String

On Error GoTo vError

If Not fxVerifica Then
  Exit Sub
End If

If vEditar Then
 If Trim(txtPlan) <> vCodigo Then
   MsgBox "Ha modificado el Código del Plan", vbExclamation
   Exit Sub
 End If
End If



If Not vEditar Then
   strSQL = "insert CRD_PLAZO_BONO(cod_Plazo_Bono,descripcion,Notas,Activo,Registro_Fecha,Registro_Usuario)" _
          & " values('" & Trim(txtPlan.Text) & "','" & txtDescripcion.Text & "','" & txtNotas.Text & "'," & chkActivo.Value _
          & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Registra", "Plazos: Plan de Bonificación : " & Trim(txtPlan))

Else
   strSQL = "update CRD_PLAZO_BONO set descripcion = '" & txtDescripcion.Text & "', Notas = '" & txtNotas.Text & "', Activo = " _
          & chkActivo.Value & " where cod_Plazo_Bono = '" & txtPlan.Text & "'"
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Modifica", "Plazos: Plan de Bonificación : " & Trim(vCodigo))

End If

Call sbLimpia(True)

vCodigo = Trim(txtPlan)
vEditar = True

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation
txtPlan.SetFocus

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
    strSQL = "insert CRD_PLAZO_BONO_ASG(cod_Plazo_Bono,garantia,registro_fecha,registro_usuario) values('" _
           & txtPlan.Text & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
    strSQL = "delete CRD_PLAZO_BONO_ASG where cod_Plazo_Bono = '" _
           & txtPlan.Text & "' and Garantia = '" & Item.Tag & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswPlanes_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Call sbConsulta(Item.Text)
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vCodigo = "" Then Exit Sub


Select Case Item.Index
  Case 1 'Tabla de Membresia
        strSQL = "select Linea,Inicio,Corte,Plazo,Registro_Usuario,Registro_Fecha" _
               & " from CRD_PLAZO_BONO_Membresia where cod_Plazo_Bono = '" & vCodigo & "'"
        Call sbCargaGrid(vGrid, 6, strSQL)
  Case 2 'Asignación de membresías
     vPaso = True
     
     lsw.ListItems.Clear
     strSQL = "select G.garantia,G.descripcion,A.registro_fecha" _
            & " from CRD_GARANTIA_TIPOS G left join CRD_PLAZO_BONO_ASG A on G.garantia = A.GARANTIA" _
            & " and  A.cod_Plazo_Bono = '" & vCodigo & "'"
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!Descripcion)
           itmX.Tag = rs!Garantia
        If Not IsNull(rs!registro_fecha) Then
           itmX.Checked = True
        End If
       
       rs.MoveNext
     Loop
     rs.Close
     vPaso = False
     
End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "nuevo"
    vEditar = False
    Call sbToolBar(Me.tlb, "edicion")
    Call sbLimpia
    txtPlan.SetFocus
    
  Case "editar"
    
    vEditar = True
    vCodigo = Trim(txtPlan)
    Call sbToolBar(tlb, "edicion")
    txtDescripcion.SetFocus
        
  Case "borrar"
    Call sbBorrar
        
  Case "guardar"
    Call sbGuardar
    
  Case "deshacer"
    vEditar = False
    Call sbToolBar(tlb, "nuevo")
    Call RefrescaTags(Me)
    Call sbLimpia
    txtPlan.SetFocus
    
  Case "consultar"
    
End Select

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkActivo.SetFocus
End Sub

Private Sub txtPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub txtPlan_LostFocus()
 Call sbConsulta(txtPlan.Text)
End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vLinea As Long

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1


If vGrid.Text = "" Then 'Insertar
  
  strSQL = "select isnull(max(LINEA),0) + 1 as Linea from CRD_PLAZO_BONO_MEMBRESIA " _
         & " where cod_Plazo_Bono = '" & txtPlan.Text & "'"
  Call OpenRecordSet(rs, strSQL)
   vLinea = rs!Linea
  rs.Close
     
  strSQL = "insert into CRD_PLAZO_BONO_MEMBRESIA(cod_Plazo_Bono,Linea,Inicio,Corte,Plazo,registro_fecha,registro_usuario) values('" _
         & vCodigo & "'," & vLinea & ","
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & ","
  vGrid.col = 3
  strSQL = strSQL & vGrid.Text & ","
  vGrid.col = 4
  strSQL = strSQL & vGrid.Text & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  vGrid.Text = CStr(vLinea)
  
  Call Bitacora("Registra", "Tasas Bonfificación: P:" & txtPlan.Text & "..L: " & vGrid.Text)
Else 'Actualizar

 vGrid.col = 2
 strSQL = "update CRD_PLAZO_BONO_MEMBRESIA set Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "', Inicio = " & vGrid.Text & ", Corte = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Text & ",Plazo = "
 vGrid.col = 4
 strSQL = strSQL & vGrid.Text & " where cod_Plazo_Bono = '" & vCodigo & "' and Linea = "
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Tasas Bonfificación: P:" & txtPlan.Text & "..L: " & vGrid.Text)

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        vGrid.col = 1
        strSQL = "delete CRD_PLAZO_BONO_MEMBRESIA where cod_Plazo_Bono = '" & txtPlan.Text & "' and Linea = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tasas Bonfificación: P:" & txtPlan.Text & "..L: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
