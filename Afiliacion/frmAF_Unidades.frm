VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Unidades 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unidades Programáticas y de Trabajo"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   HelpContextID   =   1013
   Icon            =   "frmAF_Unidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   8070
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
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8640
      Top             =   360
   End
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Unidad Programática"
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
      Value           =   -1  'True
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Realiza una consulta personalizada sobre los datos actuales"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Object.ToolTipText     =   "Imprime el listado seleccionado"
            Object.Tag             =   "1"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFUP"
                  Object.Tag             =   "1"
                  Text            =   "Lista de Unidades Programáticas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFUT"
                  Object.Tag             =   "1"
                  Text            =   "Lista de Unidades de Trabajo"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SOCRUP"
                  Object.Tag             =   "1"
                  Text            =   "Socios por Unidad Programática"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SOCUPT"
                  Object.Tag             =   "1"
                  Text            =   "Socios por Unidad Programática (Todas)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra esta ventana"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboProvincia 
      Height          =   330
      Left            =   7080
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Unidad de Trabajo"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   330
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   8655
      _Version        =   1441793
      _ExtentX        =   15266
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   330
      Left            =   8760
      TabIndex        =   11
      ToolTipText     =   "Exportar a Excel"
      Top             =   2040
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   6
      Picture         =   "frmAF_Unidades.frx":08CA
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Provincia"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Código"
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
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblConsulta 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Unidades Registradas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmAF_Unidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnEditar As Boolean
Dim mLimpiaCombo As Boolean

Dim mblnEOF As Boolean

Dim mstrCod As String

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Sub sbUnidades_List()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtCodigo.Text = ""
txtDescripcion.Text = ""

If rbTipo.Item(0).Value Then
     strSQL = "Select U.Codigo, U.Descripcion, isnull(P.Provincia,'') as 'Provincia',isnull(P.Descripcion,'') as 'ProvinciaDesc'" _
            & " from UProgramatica U left join Provincias P on U.Provincia = P.Provincia" _
            & " Where U.Descripcion like '%" & txtFiltro.Text & "%'" _
            & " Order by U.Descripcion"

Else
     strSQL = "Select U.UT_Codigo as 'Codigo', U.UT_Descripcion as 'Descripcion', '' as 'Provincia', '' as 'ProvinciaDesc'" _
            & " from UTrabajo U " _
            & " Where U.UT_Descripcion like '%" & txtFiltro.Text & "%'" _
            & " Order by U.UT_Descripcion"

End If

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!ProvinciaDesc
     itmX.SubItems(3) = rs!Provincia
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Sub Guardar()

Dim intProvincia As Integer
Dim strSQL As String

If rbTipo(0).Value Then
  If Trim(txtCodigo.Text) <> "" And Trim(txtDescripcion.Text) <> "" And _
    cboProvincia.Text <> "" Then
    
    mLimpiaCombo = True
    
    intProvincia = cboProvincia.ItemData(cboProvincia.ListIndex)
    
    If mblnEditar = True Then
       strSQL = "Update UProgramatica Set Codigo = '" & Trim(txtCodigo) & "', Descripcion = '" & Trim(txtDescripcion.Text) _
              & "', Provincia = " & intProvincia _
              & " Where Codigo = '" & txtCodigo.Text & "'"
       Call ConectionExecute(strSQL)
       Call Bitacora("Modifica", "Modifico Unidad Programatica " & Trim(txtCodigo))
    Else
       strSQL = "Insert UProgramatica (Codigo,Descripcion,Provincia)"
       strSQL = strSQL & " Values('" & Trim(txtCodigo) & "','"
       strSQL = strSQL & (Trim(txtDescripcion)) & "',"
       strSQL = strSQL & intProvincia & ")"
       Call ConectionExecute(strSQL)
       Call Bitacora("Registra", "Registro Unidad Programatica " & Trim(txtCodigo))
    End If
    
    Call sbUnidades_List
    
    tlbPrincipal.Buttons.Item(1).Enabled = True
    tlbPrincipal.Buttons.Item(2).Enabled = True
    tlbPrincipal.Buttons.Item(3).Enabled = True
    tlbPrincipal.Buttons.Item(4).Enabled = False
    tlbPrincipal.Buttons.Item(5).Enabled = False
    tlbPrincipal.Buttons.Item(6).Enabled = True
    tlbPrincipal.Buttons.Item(7).Enabled = True
    tlbPrincipal.Buttons.Item(8).Enabled = True
    tlbPrincipal.Buttons.Item(9).Enabled = True

    Call RefrescaTags(Me)
    
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
  
  Else
    MsgBox "Faltan Datos", vbExclamation, "Atencion!"
  End If
End If


If rbTipo(1).Value Then
  If Trim(txtCodigo) <> "" And Trim(txtDescripcion) <> "" Then
    If mblnEditar = True Then
       strSQL = "Update UTrabajo set UT_Codigo='"
       strSQL = strSQL & Trim(txtCodigo) & "',UT_Descripcion='"
       strSQL = strSQL & UCase(Trim(txtDescripcion)) & "'"
       strSQL = strSQL & " Where UT_Codigo='" & Trim(txtCodigo.Text) & "'"
       Call ConectionExecute(strSQL)
       Call Bitacora("Modifica", "Modifico Unidad Trabajo " & Trim(txtCodigo))
    Else
       strSQL = "Insert into UTrabajo (UT_Codigo,UT_Descripcion)"
       strSQL = strSQL & " Values('" & Trim(txtCodigo) & "','"
       strSQL = strSQL & UCase(Trim(txtDescripcion)) & "')"
       Call ConectionExecute(strSQL)
       Call Bitacora("Registra", "Registro Unidad Trabajo " & Trim(txtCodigo))
    End If
    
    
    Call sbUnidades_List
    
    
    tlbPrincipal.Buttons.Item(1).Enabled = True
    tlbPrincipal.Buttons.Item(2).Enabled = True
    tlbPrincipal.Buttons.Item(3).Enabled = True
    tlbPrincipal.Buttons.Item(4).Enabled = False
    tlbPrincipal.Buttons.Item(5).Enabled = False
    tlbPrincipal.Buttons.Item(6).Enabled = True
    tlbPrincipal.Buttons.Item(7).Enabled = True
    tlbPrincipal.Buttons.Item(8).Enabled = True
    tlbPrincipal.Buttons.Item(9).Enabled = True

    Call RefrescaTags(Me)
    
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
  Else
    MsgBox "Faltan Datos", vbExclamation, "Atencion!"
  End If
End If

End Sub

Sub Modificar()

mLimpiaCombo = False

tlbPrincipal.Buttons.Item(1).Enabled = False
tlbPrincipal.Buttons.Item(2).Enabled = False
tlbPrincipal.Buttons.Item(3).Enabled = True
tlbPrincipal.Buttons.Item(4).Enabled = True
tlbPrincipal.Buttons.Item(5).Enabled = True
tlbPrincipal.Buttons.Item(6).Enabled = False
tlbPrincipal.Buttons.Item(7).Enabled = False
tlbPrincipal.Buttons.Item(8).Enabled = False
tlbPrincipal.Buttons.Item(9).Enabled = False

Call RefrescaTags(Me)

mblnEditar = True
mLimpiaCombo = True

End Sub


Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
 
 vModulo = 1

        
End Sub


Private Sub Form_Load()

On Error GoTo vError

 Call sbToolBarIconos(tlbPrincipal, False)
 
 lsw.ColumnHeaders.Clear
 lsw.ColumnHeaders.Add , , "Código", 1000, vbCenter
 lsw.ColumnHeaders.Add , , "Descripción", 5000
 lsw.ColumnHeaders.Add , , "Provincia", 1300, vbCenter
 lsw.ColumnHeaders.Add , , "P.Id", 10, vbCenter
  

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtCodigo.Text = Item.Text
txtDescripcion.Text = Item.SubItems(1)
Call sbCboAsignaDato(cboProvincia, Item.SubItems(2), False, Item.SubItems(3))

End Sub


Private Sub rbTipo_Click(Index As Integer)
Call sbUnidades_List

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

strSQL = "select Provincia as 'IdX', Descripcion as 'ItmX' from Provincias"
Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)

Call sbUnidades_List

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strResp As String
Dim strSQL As String

On Error GoTo ErrorTransaccion


If Button.Key <> "cerrar" Then
   Me.MousePointer = vbHourglass
End If

Select Case Button.Key
    Case "insertar"
            txtCodigo.SetFocus
            
            mblnEditar = False
            mLimpiaCombo = True
            
            tlbPrincipal.Buttons.Item(1).Enabled = False
            tlbPrincipal.Buttons.Item(2).Enabled = False
            tlbPrincipal.Buttons.Item(3).Enabled = False
            tlbPrincipal.Buttons.Item(4).Enabled = True
            tlbPrincipal.Buttons.Item(5).Enabled = True
            tlbPrincipal.Buttons.Item(6).Enabled = False
            tlbPrincipal.Buttons.Item(7).Enabled = False
            tlbPrincipal.Buttons.Item(8).Enabled = False
            tlbPrincipal.Buttons.Item(9).Enabled = False
                
    Case "modificar"
         Call Modificar
                
    Case "borrar"
           If lsw.SelectedItem.Text <> "" Then
              strResp = MsgBox("Registro Será Eliminado", vbQuestion + vbYesNo, "Confirma Eliminación?")
              
              If strResp = vbYes Then
                strSQL = "Delete From " _
                        & IIf(rbTipo(0).Value = True, "UProgramatica", "UTrabajo")
                strSQL = strSQL & " Where "
                strSQL = strSQL & IIf(rbTipo(0).Value = True, "Codigo='", "UT_Codigo='")
                strSQL = strSQL & txtCodigo.Text & "'"
                
                Call ConectionExecute(strSQL)
                Call Bitacora("Borra", "Elimino " & IIf(rbTipo(0).Value = True, "Programatica", "Trabajo") & Trim(txtCodigo))
                
                txtCodigo = ""
                txtDescripcion = ""
                
                mblnEditar = False
                mLimpiaCombo = True
              
               Call sbUnidades_List
                
                If mblnEOF = True Then
                   tlbPrincipal.Buttons.Item(1).Enabled = True
                   tlbPrincipal.Buttons.Item(2).Enabled = False
                   tlbPrincipal.Buttons.Item(3).Enabled = False
                   tlbPrincipal.Buttons.Item(4).Enabled = False
                   tlbPrincipal.Buttons.Item(5).Enabled = False
                   tlbPrincipal.Buttons.Item(6).Enabled = False
                   tlbPrincipal.Buttons.Item(7).Enabled = False
                   tlbPrincipal.Buttons.Item(8).Enabled = True
                   tlbPrincipal.Buttons.Item(9).Enabled = True
                   lsw.Enabled = False
                Else
                   tlbPrincipal.Buttons.Item(1).Enabled = True
                   tlbPrincipal.Buttons.Item(2).Enabled = True
                   tlbPrincipal.Buttons.Item(3).Enabled = True
                   tlbPrincipal.Buttons.Item(4).Enabled = False
                   tlbPrincipal.Buttons.Item(5).Enabled = False
                   tlbPrincipal.Buttons.Item(6).Enabled = True
                   tlbPrincipal.Buttons.Item(7).Enabled = True
                   tlbPrincipal.Buttons.Item(8).Enabled = True
                   tlbPrincipal.Buttons.Item(9).Enabled = True
                   lsw.Enabled = True
                End If
                Call RefrescaTags(Me)
              End If
            End If

           
    Case "guardar"
         Call Guardar
    
    Case "deshacer"
            txtCodigo.Text = ""
            txtDescripcion.Text = ""

            mblnEditar = False
            mLimpiaCombo = True

            tlbPrincipal.Buttons.Item(1).Enabled = True
            tlbPrincipal.Buttons.Item(4).Enabled = False
            tlbPrincipal.Buttons.Item(5).Enabled = False
            tlbPrincipal.Buttons.Item(8).Enabled = True
            tlbPrincipal.Buttons.Item(9).Enabled = True
                
            If mblnEOF = True Then
               tlbPrincipal.Buttons.Item(2).Enabled = False
               tlbPrincipal.Buttons.Item(3).Enabled = False
               tlbPrincipal.Buttons.Item(6).Enabled = False
               tlbPrincipal.Buttons.Item(7).Enabled = False
               lsw.Enabled = False
            Else
               tlbPrincipal.Buttons.Item(2).Enabled = True
               tlbPrincipal.Buttons.Item(3).Enabled = True
               tlbPrincipal.Buttons.Item(6).Enabled = True
               tlbPrincipal.Buttons.Item(7).Enabled = True
               lsw.Enabled = True
            End If
            Call RefrescaTags(Me)


    Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

    Case "cerrar"
        UnLoad Me
        
    Case "imprimir"
    
    Case "consultar"
End Select


If Button.Key <> "cerrar" Then
   Me.MousePointer = vbDefault
End If

Exit Sub

ErrorTransaccion:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla el reporte elegido por el usuario.
'REFERENCIAS:   ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Personas"

 .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 
 .Connect = glogon.ConectRPT

Select Case ButtonMenu.Key
            
   Case "AFUP"
      .ReportFileName = App.Path & "\Reportes\AfiListaUnidadProgramatica.rpt"
   
   Case "AFUT"
      .ReportFileName = App.Path & "\Reportes\AfiListaUnidadTrabajo.rpt"
   
   Case "SOCRUP"
      .ReportFileName = App.Path & "\Reportes\AfiDetalleSociosporUnidad.rpt"
      strSQL = InputBox("Especifique el Codigo de la Unidad Programática", "Afiliaciones Para Unidad Programática")
      .SelectionFormula = "{SOCIOS.ESTADOACTUAL} ='S' And {UPROGRAMATICA.CODIGO}='" & strSQL & "'"
   
   Case "SOCUPT"
      .ReportFileName = App.Path & "\Reportes\AfiDetalleSociosporUnidad.rpt"
      .SelectionFormula = "{SOCIOS.ESTADOACTUAL} ='S'"
End Select

 .PrintReport

End With
Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
On Error GoTo vError

If KeyAscii = vbKeyReturn Then
   txtDescripcion.SetFocus
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_LostFocus()

On Error GoTo vError

Me.MousePointer = vbHourglass

If rbTipo(0).Value = True Then
   strSQL = "Select U.Codigo, U.Descripcion, P.Descripcion as 'ProvinciaDesc', isnull(P.Provincia,'') as 'Provincia'" _
          & " from UProgramatica U left join Provincias P on U.Provincia = P.Provincia Where U.Codigo = '" & Trim(txtCodigo) & "'"
Else
   strSQL = "Select UT_Codigo as 'CODIGO', UT_DESCRIPCION as 'DESCRIPCION', '' as 'ProvinciaDesc', '' as 'PROVINCIA'" _
          & " from UTrabajo Where UT_Codigo = '" & Trim(txtCodigo) & "'"
End If

Call OpenRecordSet(rs, strSQL)

If rs.EOF And rs.BOF Then
    txtDescripcion.Text = ""
    mblnEditar = False
Else
    mblnEditar = True
    txtCodigo.Text = rs!Codigo
    txtDescripcion.Text = rs!Descripcion
    If rs!Provincia <> "" Then
        Call sbCboAsignaDato(cboProvincia, rs!ProvinciaDesc, True, rs!Provincia)
    End If
End If


tlbPrincipal.Buttons.Item(4).Enabled = True
Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbUnidades_List
End If
End Sub
