VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCntX_ConAsientos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asientos de Eliminación - Consolidaciones"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtNAsiento 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   9015
   End
   Begin VB.TextBox txtDiferencia 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7656
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1692
   End
   Begin VB.TextBox txtCredito 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1692
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   8880
      TabIndex        =   4
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   37278
   End
   Begin MSComctlLib.ImageList ImageListMenu 
      Left            =   9480
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConAsientos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConAsientos.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConAsientos.frx":0D2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1005
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
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4332
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   11292
      _Version        =   524288
      _ExtentX        =   19918
      _ExtentY        =   7641
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
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
      SpreadDesigner  =   "frmCntX_ConAsientos.frx":113E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Consolidación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Num. Asiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6756
      TabIndex        =   9
      Top             =   5916
      Width           =   648
   End
   Begin VB.Label lsblr 
      Caption         =   "Diferencia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4116
      TabIndex        =   8
      Top             =   5916
      Width           =   792
   End
   Begin VB.Label lblAsientoEstado 
      Caption         =   "Estado del Asiento."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   432
      Left            =   120
      TabIndex        =   7
      Top             =   5820
      Width           =   2592
   End
End
Attribute VB_Name = "frmCntX_ConAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type xUltimos
   NumAsiento As String
   Detalle    As String
   fecha      As Date
   AplAsiento As Integer
End Type

Dim vEdita As Boolean, vBusca As Integer, vUltimos As xUltimos

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub sbLimpiezaParcial(iCodigo As Integer)

vGrid.MaxRows = 0
vGrid.MaxRows = 1

txtDescripcion = ""

Select Case iCodigo
  Case 1 'Cambia el Tipo de Asiento
    txtNAsiento = ""
  Case 2 'Cambia el periodo
    txtNAsiento = ""
   
End Select

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Set Me.Icon = frmContenedor.Icon
vPaso = False

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vEdita = False
 Call sbLimpiaPantalla
 'Call sbToolBar(tlb, "edicion")
 
 strSQL = "select * from CNTX_CONSOLIDA_DEFINICION"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 cbo.Clear
 
 Do While Not rs.EOF
   cbo.AddItem Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
   cbo.ItemData(cbo.NewIndex) = rs!COD_CONSOLIDA
   vPaso = True
   rs.MoveNext
 Loop
 
 If vPaso Then
   rs.MoveFirst
   cbo.Text = Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
 End If
 rs.Close
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub

Private Function fxVerificaAsiento() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, lng As Long

'Fecha del Asiento vrs Periodo
'Numero de Asiento
'Cuentas (En el Detalle)

fxVerificaAsiento = True
vMensaje = ""


For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 1
 If vGrid.Text <> "" Then
   vGrid.Col = 2
   If vGrid.Text = "" Then
      vGrid.Col = 1
      vMensaje = vMensaje & vbCrLf & "- Cuenta " & vGrid.Text & " No Existe"
   End If
 End If
Next lng


If CCur(txtDiferencia) <> 0 Then vMensaje = vMensaje & vbCrLf & "- El asiento no se encuentra Balanceado, por motivos" _
            & " de Seguridad no se permite Guardar Dicha información en asientos de CNTX_CONSOLIDA_DEFINICION"

If Len(vMensaje) > 0 Then
  fxVerificaAsiento = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbLimpiaPantalla()
vBusca = 1
txtCredito = 0
txtDebito = 0
txtDescripcion = ""
txtDiferencia = 0
txtNAsiento = ""
vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 6
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNAsiento.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      If vUltimos.AplAsiento = 0 Then
        vEdita = True
        txtDescripcion.SetFocus
        Call sbToolBar(tlb, "edicion")
      Else
        MsgBox "Este Asiento ya Fue Aplicado (No se puede Modificar/Borrar)", vbInformation
      End If
    
    Case "BORRAR"
      If vUltimos.AplAsiento = 0 Then
        Call sbBorrar
      Else
        MsgBox "Este Asiento ya Fue Aplicado (No se puede Modificar/Borrar)", vbInformation
      End If
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case "DESHACER"
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
    
    Case "CONSULTAR"
      Select Case vBusca
       Case 3
            gBusquedas.Columna = "descripcion"
            gBusquedas.Orden = "descripcion"
            gBusquedas.Consulta = "select cod_asiento,descripcion,fecha from con_asientos"
            gBusquedas.Filtro = " and cod_consolida = " & cbo.ItemData(cbo.ListIndex)
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            txtNAsiento.SetFocus
       Case 4
            gBusquedas.Columna = "Cod_asiento"
            gBusquedas.Orden = "Cod_asiento"
            gBusquedas.Consulta = "select cod_asiento,descripcion,fecha from con_asientos"
            gBusquedas.Filtro = " and cod_consolida = " & cbo.ItemData(cbo.ListIndex)
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            txtNAsiento.SetFocus
       End Select
    
    Case "REPORTES"
      
      strSQL = "{CON_ASIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex) _
             & " {CON_ASIENTOS.COD_ASIENTO} = '" & txtNAsiento & "'"
      
'      Call sbReportes("ASIENTO", strSQL)
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    
    Case "CERRAR"
      UnLoad Me
End Select

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    If i = 1 Then
        vGrid.Text = fxConsolida_CuentaFormato(True, CStr(rs.Fields(i - 1).Value), cbo.ItemData(cbo.ListIndex))

    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
    End If
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub sbConsultaAsiento(strNumero As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from con_asientos where cod_consolida = " & cbo.ItemData(cbo.ListIndex) _
       & " and cod_asiento = '" & strNumero & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  vUltimos.NumAsiento = rs!cod_asiento
  vUltimos.fecha = rs!fecha
  vUltimos.AplAsiento = IIf((rs!aplicado = "N"), 0, 1)
  
  'llenar datos en pantalla
  
  txtDescripcion = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
  dtpFecha.Value = vUltimos.fecha
  txtNAsiento = vUltimos.NumAsiento
  
  If vUltimos.AplAsiento = 0 Then
    lblAsientoEstado.Caption = "Este Asiento se Encuentra Pendiente"
  Else
    lblAsientoEstado.Caption = "Este Asiento se Encuentra Mayorizado"
  End If
  
strSQL = "select A.cod_cuenta,B.descripcion,detalle,debitos,creditos,linea" _
       & " from con_asientos_detalle A inner join cuentas B on A.cod_cuenta = B.cod_cuenta" _
       & " where B.COD_CONTABILIDAD in(select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION" _
       & " where cod_consolida = " & cbo.ItemData(cbo.ListIndex) & ") and cod_asiento = '" _
       & vUltimos.NumAsiento & "'" & " order by linea"
   
  Call sbCargaGridLocal(vGrid, 6, strSQL)
 
  Call sbSumaDebitosCreditos

  If vUltimos.AplAsiento = 1 Then
    vGrid.Lock = True
  Else
    vGrid.Lock = False
  End If

End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbGuardar()
Dim strSQL As String, lng As Long

On Error GoTo vError

If fxVerificaAsiento Then
      
    If vEdita Then
      
      strSQL = "update con_asientos set descripcion = '" & UCase(txtDescripcion) _
             & "',fecha = '" & Format(dtpFecha.Value, "yyyy/mm/dd") _
             & "' where cod_consolida = " & cbo.ItemData(cbo.ListIndex) _
             & " and cod_asiento = '" & txtNAsiento & "'"
      glogon.Conection.Execute strSQL
     
      strSQL = "delete con_asientos_detalle where cod_consolida = " _
             & cbo.ItemData(cbo.ListIndex) & " and cod_asiento = '" & txtNAsiento & "'"
      glogon.Conection.Execute strSQL
    
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.Col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into con_asientos_detalle(cod_asiento,cod_consolida" _
                   & ",linea,cod_cuenta,detalle,debitos,creditos" _
                   & ") values('" & txtNAsiento & "'," & cbo.ItemData(cbo.ListIndex) _
                   & "," & lng & ",'"
            vGrid.Row = lng
            vGrid.Col = 1
            strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
            vGrid.Col = 3
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.Col = 4
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.Col = 5
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")" _
          
            glogon.Conection.Execute strSQL
              
              
              
         End If 'vgrid.Text <> ""
       
       Next lng
    
      Call Bitacora("Modifica", "Asiento Consolidado: " & txtNAsiento & " Consolida:" & cbo.ItemData(cbo.ListIndex))
    
    
    Else 'Inserta
       strSQL = "insert into con_asientos(cod_consolida,cod_asiento" _
              & ",fecha,descripcion,aplicado) values(" & cbo.ItemData(cbo.ListIndex) _
              & ",'" & txtNAsiento & "','" & Format(dtpFecha.Value, "yyyy/mm/dd") & "','" _
              & UCase(txtDescripcion) & "','N')"
       glogon.Conection.Execute strSQL
       
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.Col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into con_asientos_detalle(cod_asiento,cod_consolida" _
                   & ",linea,cod_cuenta,detalle,debitos,creditos" _
                   & ") values('" & txtNAsiento & "'," & cbo.ItemData(cbo.ListIndex) _
                   & "," & lng & ",'"
            vGrid.Row = lng
            vGrid.Col = 1
            strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
            vGrid.Col = 3
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.Col = 4
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.Col = 5
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")" _
          
            glogon.Conection.Execute strSQL
              
         End If 'vgrid.Text <> ""
       
       Next lng
       
       
       Call Bitacora("Registra", "Asiento Consolidado: " & txtNAsiento & " Consolida : " & cbo.ItemData(cbo.ListIndex))
        
    End If 'Si Inserta o Actualiza

        Call sbToolBar(tlb, "activo")
        Call sbConsultaAsiento(txtNAsiento)
        
        vEdita = True
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation


End If 'Verificacion del Asiento

' Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String
On Error GoTo vError
i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
  strSQL = "delete con_asientos_detalle where cod_consolida = " _
         & cbo.ItemData(cbo.ListIndex) & " and cod_asiento = '" & txtNAsiento & "'"
  glogon.Conection.Execute strSQL
  
  strSQL = "delete con_asientos where cod_consolida = " _
         & cbo.ItemData(cbo.ListIndex) & " and cod_asiento = '" & txtNAsiento & "'"
  glogon.Conection.Execute strSQL
  

  Call Bitacora("Elimina", "Asiento Consolida: " & txtNAsiento & " Consolida:" _
                  & cbo.ItemData(cbo.ListIndex))

  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtDescripcion_GotFocus()
vBusca = 4
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtNAsiento_GotFocus()
vBusca = 3
End Sub

Private Sub txtNAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 dtpFecha.SetFocus
 Call sbConsultaAsiento(txtNAsiento)
End If
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub


Private Function fxVerificaCuenta(strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & cbo.ItemData(cbo.ListIndex)

rsX.Open strSQL, glogon.Conection, adOpenStatic
strSQL = "select isnull(count(*),0) as Existe from cuentas where COD_CONTABILIDAD = " & rsX!COD_CONTABILIDAD _
       & " and cod_cuenta = '" & strCuenta & "' and acepta_movimientos = 'S'"
rsX.Close

rsX.Open strSQL, glogon.Conection, adOpenStatic
fxVerificaCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function

Private Sub sbSumaDebitosCreditos()
Dim x As Long
  txtDebito = 0
  txtCredito = 0
  For x = 1 To vGrid.MaxRows
      vGrid.Row = x
      vGrid.Col = 4
      txtDebito = CCur(txtDebito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
      vGrid.Col = 5
      txtCredito = CCur(txtCredito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(6) As Variant, x As Integer
Dim strSQL As String, rs As New ADODB.Recordset

'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 6
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  
  
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To 6
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To 6
       vGrid.Col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
  
End If


If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 Then
  strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & cbo.ItemData(cbo.ListIndex)
  rs.Open strSQL, glogon.Conection, adOpenStatic
  
  gBusquedas.Columna = "cod_cuenta"
  gBusquedas.Orden = "cod_cuenta"
  gBusquedas.Filtro = " and acepta_movimientos = 'S' and COD_CONTABILIDAD = " & rs!COD_CONTABILIDAD
  gBusquedas.Consulta = "select cod_cuenta, descripcion from cuentas"
  
  rs.Close
  
  frmBusquedas.Show vbModal
  
'  frmConsultaCuentas.Show vbModal
'  vGrid.Col = vGrid.ActiveCol
'  vGrid.Row = vGrid.ActiveRow
'  vGrid.Text = gCuenta
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow

  vGrid.Text = gBusquedas.Resultado
End If

If (KeyCode = 13 Or KeyCode = vbKeyTab) Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxConsolida_CuentaFormato(True, vGrid.Text, cbo.ItemData(cbo.ListIndex))
        i = fxConsolida_CuentaFormato(False, vGrid.Text, cbo.ItemData(cbo.ListIndex))
        If fxVerificaCuenta(CStr(i)) Then
          vGrid.Col = 2
          vGrid.Text = fxConsolida_Cuenta("D", CStr(i), cbo.ItemData(cbo.ListIndex))
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CUENTAS", vbCritical
        End If
        
      Case 3
        vUltimos.Detalle = vGrid.Text
        
      Case 4 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 5 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
        
        End If
      
      Case 6 'Nueva linea
        If vGrid.MaxRows = vGrid.Row Then
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.MaxRows
            vGrid.Col = 3
            vGrid.Text = vUltimos.Detalle
        End If
    End Select
End If

If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 3
    vGrid.Text = vUltimos.Detalle
End If


End Sub



