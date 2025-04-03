VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmUS_Access_Estaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Accesos: Estaciones Autorizadas"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   9735
      _Version        =   1441793
      _ExtentX        =   17171
      _ExtentY        =   11245
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
      ItemCount       =   2
      Item(0).Caption =   "Estaciones"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "gbMACs"
      Item(1).Caption =   "Vincular"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "Label1"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6015
         Left            =   -67120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12091
         _ExtentY        =   10610
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox gbMACs 
         Height          =   2055
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Width           =   9135
         _Version        =   1441793
         _ExtentX        =   16113
         _ExtentY        =   3625
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Begin XtremeSuiteControls.PushButton btnMACs 
            Height          =   375
            Index           =   0
            Left            =   6120
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Guardar"
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
            Picture         =   "frmUS_Access_Estaciones.frx":0000
         End
         Begin XtremeSuiteControls.ComboBox cboMAC_1 
            Height          =   330
            Left            =   2880
            TabIndex        =   10
            Top             =   600
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboMAC_2 
            Height          =   330
            Left            =   2880
            TabIndex        =   11
            Top             =   1080
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnMACs 
            Height          =   375
            Index           =   1
            Left            =   7560
            TabIndex        =   13
            Top             =   1560
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1085
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
            Picture         =   "frmUS_Access_Estaciones.frx":0731
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   1080
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "MAC Autorizada No. 2"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   8
            Top             =   600
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "MAC Autorizada No. 1"
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
         Begin XtremeShortcutBar.ShortcutCaption scEquipo 
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   9135
            _Version        =   1441793
            _ExtentX        =   16113
            _ExtentY        =   661
            _StockProps     =   14
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
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5775
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   9375
         _Version        =   524288
         _ExtentX        =   16536
         _ExtentY        =   10186
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
         MaxCols         =   494
         ScrollBars      =   2
         SpreadDesigner  =   "frmUS_Access_Estaciones.frx":0D6F
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   2655
         Left            =   -69520
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   4683
         _StockProps     =   79
         Caption         =   "Lista de Estaciones que han accedido al cliente y que no se encuentran vinculadas"
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   240
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
            Picture         =   "frmUS_Access_Estaciones.frx":13C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Access_Estaciones.frx":1BA1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9840
      Top             =   840
   End
   Begin XtremeSuiteControls.FlatEdit txtCliente 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   9735
      _Version        =   1441793
      _ExtentX        =   17171
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _Version        =   1441793
      _ExtentX        =   17171
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Estaciones de Trabajo Autorizadas "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
Attribute VB_Name = "frmUS_Access_Estaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean, mToken As String


Private Sub btnMACs_Click(Index As Integer)


On Error GoTo vError

If Index = 0 Then
   If cboMAC_1.Text = cboMAC_2.Text And cboMAC_1.Text = "" Then
      MsgBox "La MAC No 1 y 2 son la misma, no se puede realizar el registro", vbInformation
      Exit Sub
   End If
   
   If cboMAC_1.Text = "" Then
      MsgBox "No se indicó una MAC No.1 (Principal)...verifique!", vbInformation
      Exit Sub
   End If
           
           
'spPGX_Estacion_MAC_Update (@Cliente int,  @Estacion varchar(100), @Usuario varchar(30)
'                                           , @MAC_1   varchar(100), @MAC_2    varchar(100) , @Token varchar(30)
'                                           , @App_Equipo varchar(100) = '', @App_Version varchar(20) = ''
'                                           , @App_MAC     varchar(100) = ''
'                                           , @App_Name varchar(30) = '')
           
   strSQL = "exec spPGX_Estacion_MAC_Update " & gPortal.Empresa_Id & ", '" & scEquipo.Tag & "', '" & glogon.Usuario & "', '" & cboMAC_1.Text _
                & "', '" & cboMAC_2.Text & "', '" & mToken & "', '" & glogon.Maquina & "', '" & glogon.AppVersion _
                & "', '" & glogon.Maquina & "', 'SystemSecurity'"
                
   Call ConectionExecute(strSQL, 1)
   
   MsgBox "Información Guardada Satisfactoriamente!", vbInformation
   
End If



scEquipo.Caption = ""
scEquipo.Tag = ""
gbMACs.Visible = False
vGrid.Height = 5775

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

    scEquipo.Caption = ""
    scEquipo.Tag = ""
    gbMACs.Visible = False
    vGrid.Height = 5775

End Sub

Private Sub Form_activate()
vModulo = 13
End Sub

Private Sub Form_Load()
vModulo = 13

txtCliente.Tag = gPortal.Empresa_Id
txtCliente.Text = gPortal.Empresa_Name

tcMain.Item(0).Selected = True

mToken = "q@$-&%1-mkE+1"


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbEstacionesVincula()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spPGX_Estacion_Loggin " & txtCliente.Tag & ",2"

vPaso = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Estaciones Detectadas", lsw.Width - 100
End With

With lsw.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!ESTACION)
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from PGX_CLIENTES_ESTACIONES" _
       & " where ESTACION = '" & vGrid.Text & "' and cod_empresa = " & txtCliente.Tag
Call OpenRecordSet(rs, strSQL, 1)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "Insert PGX_CLIENTES_ESTACIONES(COD_EMPRESA,ESTACION,DESCRIPCION, ACTIVA,Registro_Usuario,Registro_Fecha)" _
        & " values(" & txtCliente.Tag & ",'" & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',Getdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Estación de Acceso: " & vGrid.Text & " - Cliente id: " & txtCliente.Tag)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update PGX_CLIENTES_ESTACIONES set descripcion = '" & vGrid.Text & "', Activa = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where ESTACION = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "' and cod_empresa = " & txtCliente.Tag
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", " Estación de Acceso: " & vGrid.Text & " - Cliente id: " & txtCliente.Tag)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If Item.Checked Then
    strSQL = "exec spPGX_Estacion_Vincula " & txtCliente.Tag & ",'" & Item.Text & "','" & glogon.Usuario & "',1"
Else
    strSQL = "exec spPGX_Estacion_Vincula " & txtCliente.Tag & ",'" & Item.Text & "','" & glogon.Usuario & "',0"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbEstaciones_Consulta()
        
        strSQL = "select ESTACION,descripcion, Activa, 0" _
                & " from PGX_CLIENTES_ESTACIONES where cod_empresa = " & txtCliente.Tag _
                & " order by ESTACION"
        Call sbCargaGrid(vGrid, 4, strSQL)
        
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Estaciones
    Call sbEstaciones_Consulta
  
  Case 1 'Vinculacion
    Call sbEstacionesVincula
    
End Select


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


tcMain.Item(0).Selected = True
Call sbEstaciones_Consulta
End Sub


Private Sub sbMACs_Consulta(pEstacion As String, pName As String)

On Error GoTo vError

gbMACs.Visible = True

vGrid.Height = 3735

scEquipo.Caption = "Estación: " & pEstacion & " ¦ " & pName
scEquipo.Tag = pEstacion

cboMAC_1.Clear
cboMAC_2.Clear

strSQL = "exec spPGX_Estacion_Loggin_MACs " & gPortal.Empresa_Id & ", '" & pEstacion & "', '" & mToken & "'"
Call OpenRecordSet(rs, strSQL, 1)
Do While Not rs.EOF
 cboMAC_1.AddItem rs!Mac
 cboMAC_2.AddItem rs!Mac
 rs.MoveNext
Loop
rs.Close

strSQL = "select isnull(MAC_01,'') as 'MAC_01', isnull(MAC_02,'') as 'MAC_02' " _
       & "  From PGX_CLIENTES_ESTACIONES" _
       & " Where COD_EMPRESA = " & gPortal.Empresa_Id & " and ESTACION = '" & pEstacion & "'"
Call OpenRecordSet(rs, strSQL, 1)
If Not rs.EOF And Not rs.BOF Then
    If rs!Mac_01 <> "" Then
        Call sbCboAsignaDato(cboMAC_1, rs!Mac_01, True, rs!Mac_01)
    End If
    
    If rs!Mac_02 <> "" Then
        Call sbCboAsignaDato(cboMAC_2, rs!Mac_02, True, rs!Mac_02)
    End If
End If
rs.Close

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim pEstacion As String, pName As String


pEstacion = ""
pName = ""

With vGrid
  .Row = Row
  .Col = 1
  If .Text <> "" Then
    pEstacion = .Text
    .Col = 2
    pName = .Text
    
    Call sbMACs_Consulta(pEstacion, pName)
  End If
End With

End Sub

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
        vGrid.Col = 1
        strSQL = "delete PGX_CLIENTES_ESTACIONES where ESTACION = '" & vGrid.Text _
                & "' and cod_empresa = " & txtCliente.Tag
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Estación de Acceso: " & vGrid.Text & " Cliente Id: " & txtCliente.Tag)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

