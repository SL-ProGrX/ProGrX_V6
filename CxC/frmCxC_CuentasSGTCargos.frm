VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCxC_CuentasSGTCargos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Cargos de Activaci�n de la Operaci�n"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   10575
      _Version        =   1310723
      _ExtentX        =   18653
      _ExtentY        =   4260
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
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8520
      Top             =   240
   End
   Begin XtremeSuiteControls.GroupBox gbRegistro 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   10575
      _Version        =   1310723
      _ExtentX        =   18653
      _ExtentY        =   6800
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtCargoCod 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   720
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   255
         Left            =   9240
         TabIndex        =   4
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtCargoDesc 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   720
         Width           =   5775
         _Version        =   1310723
         _ExtentX        =   10186
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtValor 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   7440
         TabIndex        =   8
         Top             =   1200
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoCargo 
         Height          =   315
         Left            =   7440
         TabIndex        =   9
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnNuevo 
         Height          =   495
         Left            =   4680
         TabIndex        =   10
         Top             =   3120
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Nuevo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCxC_CuentasSGTCargos.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBorrar 
         Height          =   495
         Left            =   6120
         TabIndex        =   11
         Top             =   3120
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Borrar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCxC_CuentasSGTCargos.frx":0632
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   7680
         TabIndex        =   12
         Top             =   3120
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Guardar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCxC_CuentasSGTCargos.frx":0BD6
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   915
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   7455
         _Version        =   1310723
         _ExtentX        =   13150
         _ExtentY        =   1614
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   10575
         _Version        =   1310723
         _ExtentX        =   18653
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Registro de Cargos"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cargo"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Valor"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   15
         Top             =   1200
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto de la Operaci�n "
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   14
         Top             =   1560
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto del Cargo"
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
      End
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmCxC_CuentasSGTCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMonto As Currency, mRebajosTotales As Currency, mOperacion As Long, mCedula As String
Dim vPaso As Boolean, vScroll As Boolean, mIngresosTotales As Currency
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub btnBorrar_Click()

On Error GoTo vError

If txtCargoCod.Text = "" Then
  MsgBox "No se ha especificado ning�n cargo!", vbExclamation
  Exit Sub
End If

strSQL = "delete CxC_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & " and cod_cargo = '" _
       & txtCargoCod.Text & "'"
Call ConectionExecute(strSQL)
Call Timer1_Timer
        

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnGuardar_Click()

On Error GoTo vError

If txtCargoCod.Text = "" Then
  MsgBox "No se ha especificado ning�n cargo!", vbExclamation
  Exit Sub
End If

strSQL = "select count(*) as Existe from Cxc_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & " and cod_Cargo = '" & txtCargoCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    strSQL = "insert CxC_Cuentas_Rebajos_Cargos(cod_cargo,Operacion,tipo,monto,valor,modifica,detalle,registro_usuario" _
           & ",registro_fecha) values('" & txtCargoCod.Text & "'," & mOperacion & ",'" & Mid(cboTipo.Text, 1, 1) _
           & "'," & CCur(txtMontoCargo.Text) & "," & CDbl(txtValor.Text) & ",1,'" _
           & txtDetalle.Text & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   
Else
   strSQL = "update CxC_Cuentas_Rebajos_Cargos set tipo = '" & Mid(cboTipo.Text, 1, 1) & "', Valor = " & CDbl(txtValor.Text) _
          & ",Monto = " & CCur(txtMontoCargo.Text) & ", Detalle = '" & txtDetalle.Text & "',registro_fecha = dbo.MyGetdate()" _
          & ",registro_usuario ='" & glogon.Usuario _
          & "' where operacion = " & mOperacion & " and cod_cargo = '" & txtCargoCod.Text & "'"
End If
rs.Close

Call ConectionExecute(strSQL)
Call Timer1_Timer

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub btnNuevo_Click()

txtCargoCod.Text = ""
txtCargoDesc.Text = ""
txtValor.Text = 0
txtMonto.Text = Format(mMonto, "Standard")
txtMontoCargo.Text = 0
txtDetalle = ""
cboTipo.Text = "Monto"

End Sub

Private Sub cboTipo_Click()
txtValor.Text = 0
txtMontoCargo.Text = 0
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_cargo,descripcion from CxC_Cargos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_cargo > '" & txtCargoCod.Text & "' and Activo = 1 and Tipo = 'C'" _
              & " and cod_cargo not in(select cod_cargo from Cxc_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & ")" _
              & " order by cod_cargo asc"
    Else
       strSQL = strSQL & " where cod_cargo < '" & txtCargoCod.Text & "' and Activo = 1 and Tipo = 'C'" _
              & " and cod_cargo not in(select cod_cargo from Cxc_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & ")" _
              & " order by cod_cargo desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCargoCod.Text = rs!COD_CARGO
      txtCargoDesc.Text = rs!Descripcion
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

mOperacion = GLOBALES.gTag
Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

strSQL = "Select isnull(dbo.fxCxC_CuentaRebajos(" & mOperacion & ",'TOT'),0) as 'Rebajos', Monto,cedula" _
       & ", isnull(dbo.fxCxC_CuentaIngresos(" & mOperacion & "),0) as 'Ingresos'" _
       & " from CxC_Cuentas Where Operacion = " & mOperacion
Call OpenRecordSet(rs, strSQL)
   mRebajosTotales = rs!Rebajos
   mIngresosTotales = rs!Ingresos
   mMonto = rs!Monto
   mCedula = Trim(rs!Cedula)
rs.Close


vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
vPaso = True
    cboTipo.Clear
    cboTipo.AddItem "Porcentual"
    cboTipo.AddItem "Monto"
vPaso = False


''Si esta anulada o formalizada, no permitir modificaciones
'If Operacion.EstadoSolicitud = "N" Or Operacion.EstadoSolicitud = "F" Then
'   lsw.Enabled = False
'End If

End Sub


Private Sub optTipo_Click(Index As Integer)
Call Timer1_Timer
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

If lsw.ListItems.Count <= 0 Then Exit Sub

With Item
 If CLng(.SubItems(5)) = 1 Then
     
    txtCargoCod.Text = .Text
    txtCargoDesc.Text = .SubItems(1)
    cboTipo.Text = .SubItems(3)
    txtValor.Text = .SubItems(4)
    txtDetalle.Text = .SubItems(6)
    txtMonto.Text = Format(mMonto, "Standard")
    
    If Mid(cboTipo.Text, 1, 1) = "M" Then
       txtMontoCargo.Text = txtValor.Text
    Else
       txtMontoCargo.Text = Format(mMonto * CCur(txtValor) / 100, "Standard")
    End If
 End If
End With

End Sub


Private Sub Timer1_Timer()

Me.MousePointer = vbHourglass

Timer1.Interval = 0


vPaso = True

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

Me.Caption = "Cuentas: Rebajos de Cargos"
lblX.Caption = "Cargos...:"

'Cargos Manuales
txtCargoCod.Text = ""
txtCargoDesc.Text = ""
txtValor.Text = 0
txtMonto.Text = Format(mMonto, "Standard")
txtMontoCargo.Text = 0
cboTipo.Text = "Monto"
txtDetalle.Text = ""

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Codigo", 900
lsw.ColumnHeaders.Add , , "Descripci�n", 2900
lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Tipo", 1200
lsw.ColumnHeaders.Add , , "Valor", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Modifica", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Detalle", 1200
lsw.ColumnHeaders.Add , , "Reg.Usuario", 1200
lsw.ColumnHeaders.Add , , "Reg.Fecha", 1400

'Cargos Asignados
strSQL = "select Reb.*,Car.Descripcion" _
       & " from CxC_Cargos Car inner join CxC_Cuentas_Rebajos_Cargos Reb on Car.cod_cargo = Reb.cod_cargo" _
       & " where Reb.Operacion = " & mOperacion
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!COD_CARGO)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = IIf((rs!Tipo = "P"), "Porcentual", "Monto")
      itmX.SubItems(4) = Format(rs!Valor, "Standard")
      itmX.SubItems(5) = rs!Modifica
      itmX.SubItems(6) = rs!Detalle
      itmX.SubItems(7) = rs!Registro_Usuario
      itmX.SubItems(8) = rs!Registro_Fecha
      
  rs.MoveNext
Loop
rs.Close

vPaso = False
Me.MousePointer = vbDefault

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

Select Case Button.Key
    Case "nuevo"
        txtCargoCod.Text = ""
        txtCargoDesc.Text = ""
        txtValor.Text = 0
        txtMonto.Text = Format(mMonto, "Standard")
        txtMontoCargo.Text = 0
        txtDetalle = ""
        cboTipo.Text = "Monto"
    
    Case "borrar"
    
        If txtCargoCod.Text = "" Then
          MsgBox "No se ha especificado ning�n cargo!", vbExclamation
          Exit Sub
        End If
        
        strSQL = "delete CxC_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & " and cod_cargo = '" _
               & txtCargoCod.Text & "'"
        Call ConectionExecute(strSQL)
        Call Timer1_Timer
    
    Case "guardar"
        If txtCargoCod.Text = "" Then
          MsgBox "No se ha especificado ning�n cargo!", vbExclamation
          Exit Sub
        End If
        
        strSQL = "select count(*) as Existe from Cxc_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & " and cod_Cargo = '" & txtCargoCod.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then
            strSQL = "insert CxC_Cuentas_Rebajos_Cargos(cod_cargo,Operacion,tipo,monto,valor,modifica,detalle,registro_usuario" _
                   & ",registro_fecha) values('" & txtCargoCod.Text & "'," & mOperacion & ",'" & Mid(cboTipo.Text, 1, 1) _
                   & "'," & CCur(txtMontoCargo.Text) & "," & CDbl(txtValor.Text) & ",1,'" _
                   & txtDetalle.Text & "','" & glogon.Usuario & "',dbo.MyGetdate())"
           
        Else
           strSQL = "update CxC_Cuentas_Rebajos_Cargos set tipo = '" & Mid(cboTipo.Text, 1, 1) & "', Valor = " & CDbl(txtValor.Text) _
                  & ",Monto = " & CCur(txtMontoCargo.Text) & ", Detalle = '" & txtDetalle.Text & "',registro_fecha = dbo.MyGetdate()" _
                  & ",registro_usuario ='" & glogon.Usuario _
                  & "' where operacion = " & mOperacion & " and cod_cargo = '" & txtCargoCod.Text & "'"
        End If
        rs.Close
        
        Call ConectionExecute(strSQL)
        Call Timer1_Timer
End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtCargoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Cod_Cargo"
  gBusquedas.Orden = "Cod_Cargo"
  gBusquedas.Consulta = "select Cod_Cargo,Descripcion from CxC_Cargos"
  gBusquedas.Filtro = " and Tipo = 'C' and cod_cargo not in(select cod_cargo from Cxc_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & ")"
  frmBusquedas.Show vbModal
  txtCargoCod.Text = gBusquedas.Resultado
  txtCargoDesc.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select Cod_Cargo,Descripcion from CxC_Cargos"
  gBusquedas.Filtro = " and Tipo = 'C' and cod_cargo not in(select cod_cargo from Cxc_Cuentas_Rebajos_Cargos where Operacion = " & mOperacion & ")"
  frmBusquedas.Show vbModal
  txtCargoCod.Text = gBusquedas.Resultado
  txtCargoDesc.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtValor_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If Mid(cboTipo.Text, 1, 1) = "M" Then
   txtMontoCargo.Text = txtValor.Text
Else
   txtMontoCargo.Text = Format(mMonto * CCur(txtValor) / 100, "Standard")
End If
 
vError:
End Sub
