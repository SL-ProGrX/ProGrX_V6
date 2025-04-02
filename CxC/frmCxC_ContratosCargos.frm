VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCxC_ContratosCargos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargos Contractuales"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswC 
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   9015
      _Version        =   1310723
      _ExtentX        =   15901
      _ExtentY        =   5530
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
   End
   Begin XtremeSuiteControls.CheckBox chkModifica 
      Height          =   855
      Left            =   6120
      TabIndex        =   17
      Top             =   1440
      Width           =   3015
      _Version        =   1310723
      _ExtentX        =   5318
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Este cargo puede ser modificado por el usuario en el registro de cada Operación de cuentas x cobrar ?"
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
      Appearance      =   17
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8640
      TabIndex        =   1
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   1812
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   4473924
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboFrecuencia 
      Height          =   312
      Left            =   1200
      TabIndex        =   9
      Top             =   1920
      Width           =   1812
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   4473924
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   480
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      Top             =   480
      Width           =   5535
      _Version        =   1310723
      _ExtentX        =   9763
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCargoCod 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Top             =   960
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCargoDesc 
      Height          =   315
      Left            =   3000
      TabIndex        =   14
      Top             =   960
      Width           =   5535
      _Version        =   1310723
      _ExtentX        =   9763
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtValor 
      Height          =   315
      Left            =   4200
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2990
      _ExtentY        =   556
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDias 
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      Top             =   1920
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2990
      _ExtentY        =   556
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Frecuencia"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Días"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
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
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmCxC_ContratosCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean

Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, itmX As ListViewItem


Private Sub sbCargaLista()

Me.MousePointer = vbHourglass

vPaso = True
lswC.ListItems.Clear
strSQL = " select C.descripcion,S.*" _
       & " from CxC_Cargos C inner join CxC_Contratos_Cargos S on C.cod_cargo = S.cod_cargo" _
       & " where S.cod_contrato = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswC.ListItems.Add(, , rs!COD_CARGO)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(3) = Format(rs!Valor, "Standard")

    If rs!Tipo = "P" Then
       itmX.SubItems(2) = "Porcentual"
    Else
       itmX.SubItems(2) = "Monto"
    End If
   
   
    If rs!Frecuencia_Tipo = "O" Then
       itmX.SubItems(4) = "Operación"
    Else
       itmX.SubItems(4) = "Días"
    End If
   
    itmX.SubItems(5) = rs!Frecuencia_dias
    itmX.SubItems(6) = IIf((rs!Modifica = 1), "Sí", "No")
    itmX.SubItems(7) = rs!Registro_Usuario & "..." & rs!Registro_Fecha & ""

  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbLimpia()

vPaso = True

txtCargoCod.Text = ""
txtCargoDesc.Text = ""
txtValor.Text = "0.00"

chkModifica.Value = vbUnchecked

cboTipo.Text = "Monto"
cboFrecuencia.Text = "Operación"

vPaso = False

cboFrecuencia_Click

End Sub



Private Sub cboFrecuencia_Click()
If vPaso Then Exit Sub

If cboFrecuencia.Text = "Operación" Then
   txtDias.Enabled = False
   txtDias.Text = 0
Else
   txtDias.Enabled = True
   txtDias.Text = 30
End If

End Sub


Private Sub cboFrecuencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtDias.Enabled Then
    txtDias.SetFocus
  Else
    chkModifica.SetFocus
  End If
End If
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValor.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_cargo from CxC_Cargos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Tipo = 'C' and cod_cargo > '" & txtCargoCod.Text & "' order by cod_cargo asc"
    Else
       strSQL = strSQL & " where Tipo = 'C' and cod_cargo < '" & txtCargoCod.Text & "' order by cod_cargo desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCargoCod.Text = rs!COD_CARGO
      Call sbConsulta(txtCargoCod.Text)
    End If
End If

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()

vModulo = 31

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
vPaso = True
    cboTipo.Clear
    cboTipo.AddItem "Porcentual"
    cboTipo.AddItem "Monto"
    
    cboFrecuencia.Clear
    cboFrecuencia.AddItem "Operación"
    cboFrecuencia.AddItem "Días"
vPaso = False
 
txtCodigo.Text = GLOBALES.gTag
txtDescripcion.Text = GLOBALES.gTag2

vCodigo = txtCodigo.Text

vEdita = True

Call sbToolBarIconos(Me.tlb)
Call sbToolBar(Me.tlb, "nuevo")

With lswC.ColumnHeaders
 .Clear
 .Add , , "Código", 1300
 .Add , , "Descripción", 3500
 .Add , , "Tipo", 1400, vbCenter
 .Add , , "Valor", 1400, vbRightJustify
 .Add , , "Frecuencia", 1400
 .Add , , "Días", 1400, vbCenter
 .Add , , "Modifica", 1200, vbCenter
 .Add , , "Registro", 3000
End With

Call sbLimpia

Call sbCargaLista

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbConsulta(pCargo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

strSQL = "select P.Descripcion,C.* " _
       & " from CxC_Cargos P inner join CxC_Contratos_Cargos C on P.Cod_Cargo = C.Cod_Cargo" _
       & " where C.Cod_Cargo = '" & pCargo & "' and C.cod_contrato = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(Me.tlb, "activo")
  vEdita = True
  vCodigo = Trim(rs!COD_CARGO)
  txtCargoCod = Trim(rs!COD_CARGO)
  txtCargoDesc.Text = Trim(rs!Descripcion)
   
  If rs!Tipo = "P" Then
     cboTipo.Text = "Porcentual"
  Else
     cboTipo.Text = "Monto"
  End If
   
   
  If rs!Frecuencia_Tipo = "O" Then
     cboFrecuencia.Text = "Operación"
  Else
     cboFrecuencia.Text = "Días"
  End If
   
  txtDias.Text = rs!Frecuencia_dias
  txtValor.Text = Format(rs!Valor, "Standard")

  chkModifica.Value = rs!Modifica

Else
  'Busca Datos del Cargo Unicamente
    rs.Close
    strSQL = "select  cod_cargo,Descripcion from CxC_Cargos" _
           & " where Tipo = 'C' and Cod_Cargo = '" & pCargo & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.BOF And Not rs.EOF Then
      txtCargoCod = Trim(rs!COD_CARGO)
      txtCargoDesc.Text = rs!Descripcion
    End If
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbGuardar()

On Error GoTo vError


If vEdita Then
  strSQL = "Update CxC_Contratos_Cargos Set Valor = " & CCur(txtValor.Text) & ", frecuencia_dias = " & CLng(txtDias.Text) _
         & ",Tipo = '" & Mid(cboTipo.Text, 1, 1) & "', Frecuencia_Tipo = '" & Mid(cboFrecuencia.Text, 1, 1) _
         & "',modifica = " & chkModifica.Value _
         & " where Cod_Cargo = '" & vCodigo & "' and cod_contrato = '" & txtCodigo.Text & "'"
  Call ConectionExecute(strSQL)
  
  
  Call Bitacora("Modifica", "Cargo Suscripción Cod:" & vCodigo & " Cnt: " & Trim(txtCodigo.Text))

Else
   strSQL = "insert CxC_Contratos_Cargos(Cod_Cargo,cod_contrato,Tipo,Valor,frecuencia_Tipo,frecuencia_dias,modifica,registro_fecha,registro_usuario)" _
          & " values('" & txtCargoCod.Text & "','" & Trim(txtCodigo.Text) & "','" & Mid(cboTipo.Text, 1, 1) & "'," & CCur(txtValor.Text) _
          & ",'" & Mid(cboFrecuencia.Text, 1, 1) & "'," & CLng(txtDias.Text) & "," & chkModifica.Value & ",dbo.MyGetdate(),'" _
          & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
  
   Call Bitacora("Registra", "Cargo Suscripción Cod:" & vCodigo & " Cnt: " & Trim(txtCodigo.Text))
   
End If

vCodigo = Trim(txtCodigo)

vEdita = True

Call sbToolBar(Me.tlb, "activo")

txtCargoDesc.SetFocus

Call sbCargaLista

MsgBox "Información guardada satisfactoriamente...", vbInformation


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CxC_Contratos_Cargos where Cod_Cargo = '" & vCodigo & "' and cod_contrato = '" & txtCodigo.Text & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Borra", "Cargo Suscripción Cod:" & vCodigo & " Cnt: " & Trim(txtCodigo.Text))
  
  Call sbLimpia
  Call sbToolBar(Me.tlb, "nuevo")
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub lswC_Click()
If vPaso Then Exit Sub

With lswC.SelectedItem
  txtCargoCod.Text = .Text
  txtCargoDesc.Text = .SubItems(1)
  cboTipo.Text = .SubItems(2)
  txtValor.Text = .SubItems(3)
  cboFrecuencia.Text = .SubItems(4)
  txtDias.Text = .SubItems(5)
  chkModifica.Value = IIf(Mid(.SubItems(6), 1, 1) = "S", 1, 0)

  Call sbToolBar(Me.tlb, "activo")
End With

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpia
      Call sbToolBar(Me.tlb, "edicion")
      vCodigo = ""
      txtCargoCod.Text = ""
      txtCargoCod.SetFocus

    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCargoDesc.SetFocus
      Call sbToolBar(Me.tlb, "edicion")
      
    Case "BORRAR"
      Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
      
    Case "DESHACER"
      Call sbToolBar(Me.tlb, "nuevo")
      Call sbLimpia
      txtCargoCod.SetFocus
      vEdita = True
      
    Case "CONSULTAR"
     
    Case "REPORTES"

    Case "CERRAR"
        Unload Me
End Select

End Sub



Private Sub txtCargoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Cod_Cargo"
  gBusquedas.Orden = "Cod_Cargo"
  gBusquedas.Consulta = "select Cod_Cargo,Descripcion from CxC_Cargos"
  gBusquedas.Filtro = " and Tipo = 'C'"
  frmBusquedas.Show vbModal
  txtCargoCod.Text = gBusquedas.Resultado
  txtCargoDesc.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCargoCod_LostFocus()

If Trim(txtCargoCod.Text) <> "" Then
   Call sbConsulta(Trim(txtCargoCod))
End If

End Sub

Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select Cod_Cargo,Descripcion from CxC_Cargos"
  gBusquedas.Filtro = " and Tipo = 'C'"
  frmBusquedas.Show vbModal
  txtCargoCod.Text = gBusquedas.Resultado
  txtCargoDesc.Text = gBusquedas.Resultado2
  txtCargoCod.SetFocus
End If

End Sub


Private Sub txtDias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkModifica.SetFocus
End Sub

Private Sub txtValor_GotFocus()
On Error GoTo vError
    txtValor.Text = CCur(txtValor.Text)
vError:
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboFrecuencia.SetFocus
End Sub

Private Sub txtValor_LostFocus()
On Error GoTo vError
    txtValor.Text = Format(CCur(txtValor.Text), "Standard")
vError:

End Sub
