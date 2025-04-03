VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCxC_CuentasSGTIngresos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Otros Ingresos"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   1440
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
   Begin XtremeSuiteControls.GroupBox gbRegistro 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   10575
      _Version        =   1310723
      _ExtentX        =   18653
      _ExtentY        =   6800
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtIngresoCod 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
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
         TabIndex        =   11
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtIngresoDesc 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
      Begin XtremeSuiteControls.FlatEdit txtMontoIngreso 
         Height          =   315
         Left            =   7440
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
         TabIndex        =   19
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
         Picture         =   "frmCxC_CuentasSGTIngresos.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBorrar 
         Height          =   495
         Left            =   6120
         TabIndex        =   20
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
         Picture         =   "frmCxC_CuentasSGTIngresos.frx":0632
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   7680
         TabIndex        =   21
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
         Picture         =   "frmCxC_CuentasSGTIngresos.frx":0BD6
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   915
         Left            =   1680
         TabIndex        =   18
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   17
         Top             =   1560
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto del Ingreso"
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
         TabIndex        =   16
         Top             =   1200
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto de la Operación "
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
         TabIndex        =   8
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
         Index           =   2
         Left            =   360
         TabIndex        =   7
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
         Index           =   1
         Left            =   360
         TabIndex        =   6
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
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ingreso"
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   10575
         _Version        =   1310723
         _ExtentX        =   18653
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Registro de Ingresos"
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
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8640
      Top             =   240
   End
   Begin XtremeSuiteControls.PushButton btnActualizar 
      Height          =   612
      Left            =   7920
      TabIndex        =   1
      Top             =   240
      Width           =   2772
      _Version        =   1310723
      _ExtentX        =   4890
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Actualizar Ingresos por Comisión Repuesta"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
      Picture         =   "frmCxC_CuentasSGTIngresos.frx":1307
      ImageAlignment  =   4
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
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmCxC_CuentasSGTIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMonto As Currency, mIngresosTotales As Currency, mOperacion As Long, mCedula As String
Dim vPaso As Boolean, vScroll As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub btnActualizar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCxC_CuentaIngresoReposicion " & mOperacion & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault

Call Timer1_Timer

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnBorrar_Click()

On Error GoTo vError

        If txtIngresoCod.Text = "" Or txtIngresoCod.Tag = "" Then
          MsgBox "No se ha especificado ningún cargo!", vbExclamation
          Exit Sub
        End If
        
        strSQL = "delete CxC_Cuentas_Ingresos where Operacion = " & mOperacion & " and cod_cargo = '" _
               & txtIngresoCod.Text & "' and Linea = " & txtIngresoCod.Tag
        Call ConectionExecute(strSQL)
        Call Timer1_Timer


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnGuardar_Click()

On Error GoTo vError

        If txtIngresoCod.Text = "" Then
          MsgBox "No se ha especificado ningún cargo!", vbExclamation
          Exit Sub
        End If
        
        If txtIngresoCod.Tag = "" Then
            strSQL = "select isnull(max(Linea),0) + 1 as Linea from CxC_Cuentas_Ingresos" _
                   & " where Operacion = " & mOperacion
            Call OpenRecordSet(rs, strSQL)
               txtIngresoCod.Tag = rs!Linea
            rs.Close

            strSQL = "insert CxC_Cuentas_Ingresos(Linea,cod_cargo,Operacion,tipo,monto,valor,modifica,detalle,registro_usuario" _
                   & ",registro_fecha,cod_unidad,cod_centro_Costo) values(" & txtIngresoCod.Tag & ",'" & txtIngresoCod.Text & "'," & mOperacion & ",'" & Mid(cboTipo.Text, 1, 1) _
                   & "'," & CCur(txtMontoIngreso.Text) & "," & CDbl(txtValor.Text) & ",1,'" _
                   & txtDetalle.Text & "','" & glogon.Usuario & "',dbo.MyGetdate(),'" & GLOBALES.gOficinaUnidad & "','" & GLOBALES.gOficinaCentroCosto & "')"
           
        Else
           strSQL = "update CxC_Cuentas_Ingresos set tipo = '" & Mid(cboTipo.Text, 1, 1) & "', Valor = " & CDbl(txtValor.Text) _
                  & ",Monto = " & CCur(txtMontoIngreso.Text) & ", Detalle = '" & txtDetalle.Text & "',registro_fecha = dbo.MyGetdate()" _
                  & ",registro_usuario ='" & glogon.Usuario & "', cod_cargo = '" & txtIngresoCod.Text _
                  & "' where operacion = " & mOperacion & " and Linea = " & txtIngresoCod.Tag
        End If
        
        Call ConectionExecute(strSQL)
        Call Timer1_Timer



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnNuevo_Click()
        txtIngresoCod.Text = ""
        txtIngresoCod.Tag = ""
        txtIngresoDesc.Text = ""
        txtValor.Text = 0
        txtMonto.Text = Format(mMonto, "Standard")
        txtMontoIngreso.Text = 0
        txtDetalle = ""
        cboTipo.Text = "Monto"
End Sub

Private Sub cboTipo_Click()
txtValor.Text = 0
txtMontoIngreso.Text = 0
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_cargo,descripcion from CxC_Cargos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_cargo > '" & txtIngresoCod.Text & "' and Activo = 1 and Tipo = 'I'" _
              & " and cod_cargo not in(select cod_cargo from CxC_Cuentas_Ingresos where Operacion = " & mOperacion & ")" _
              & " order by cod_cargo asc"
    Else
       strSQL = strSQL & " where cod_cargo < '" & txtIngresoCod.Text & "' and Activo = 1 and Tipo = 'I'" _
              & " and cod_cargo not in(select cod_cargo from CxC_Cuentas_Ingresos where Operacion = " & mOperacion & ")" _
              & " order by cod_cargo desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtIngresoCod.Text = rs!COD_CARGO
      txtIngresoDesc.Text = rs!Descripcion
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

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

mOperacion = GLOBALES.gTag

strSQL = "Select isnull(dbo.fxCxC_CuentaRebajos(" & mOperacion & ",'TOT'),0) as 'Rebajos', Monto,cedula" _
       & " from CxC_Cuentas Where Operacion = " & mOperacion
Call OpenRecordSet(rs, strSQL)
   mIngresosTotales = rs!Rebajos
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


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub
If lsw.ListItems.Count <= 0 Then Exit Sub

With Item
 
 If CLng(.SubItems(5)) = 1 Then
     
    txtIngresoCod.Text = .Text
    txtIngresoCod.Tag = .Tag
    txtIngresoDesc.Text = .SubItems(1)
    cboTipo.Text = .SubItems(3)
    txtValor.Text = .SubItems(4)
    txtDetalle.Text = .SubItems(6)
    txtMonto.Text = Format(mMonto, "Standard")
    
    If Mid(cboTipo.Text, 1, 1) = "M" Then
       txtMontoIngreso.Text = txtValor.Text
    Else
       txtMontoIngreso.Text = Format(mMonto * CCur(txtValor) / 100, "Standard")
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

Me.Caption = "Cuentas: Registro de Ingresos"
lblX.Caption = "Ingresos...:"

'Cargos Manuales
txtIngresoCod.Text = ""
txtIngresoDesc.Text = ""
txtValor.Text = 0
txtMonto.Text = Format(mMonto, "Standard")
txtMontoIngreso.Text = 0
cboTipo.Text = "Monto"
txtDetalle.Text = ""


lsw.ColumnHeaders.Add , , "Codigo", 900
lsw.ColumnHeaders.Add , , "Descripción", 2900
lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Tipo", 1200
lsw.ColumnHeaders.Add , , "Valor", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Modifica", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Detalle", 1200
lsw.ColumnHeaders.Add , , "Reg.Usuario", 1200
lsw.ColumnHeaders.Add , , "Reg.Fecha", 1400

'Cargos Asignados
strSQL = "select Reb.*,Car.Descripcion" _
       & " from CxC_Cargos Car inner join CxC_Cuentas_Ingresos Reb on Car.cod_cargo = Reb.cod_cargo" _
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
      itmX.Tag = rs!Linea
  rs.MoveNext
Loop
rs.Close

vPaso = False
Me.MousePointer = vbDefault

End Sub




Private Sub tlbActualizar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCxC_CuentaIngresoReposicion " & mOperacion & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault

Call Timer1_Timer

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub txtIngresoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIngresoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Cod_Cargo"
  gBusquedas.Orden = "Cod_Cargo"
  gBusquedas.Consulta = "select Cod_Cargo,Descripcion from CxC_Cargos"
  gBusquedas.Filtro = " and Tipo = 'I' and cod_cargo not in(select cod_cargo from CxC_Cuentas_Ingresos where Operacion = " & mOperacion & ")"
  frmBusquedas.Show vbModal
  txtIngresoCod.Text = gBusquedas.Resultado
  txtIngresoDesc.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtIngresoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select Cod_Cargo,Descripcion from CxC_Cargos"
  gBusquedas.Filtro = " and Tipo = 'I' and cod_cargo not in(select cod_cargo from CxC_Cuentas_Ingresos where Operacion = " & mOperacion & ")"
  frmBusquedas.Show vbModal
  txtIngresoCod.Text = gBusquedas.Resultado
  txtIngresoDesc.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtValor_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If Mid(cboTipo.Text, 1, 1) = "M" Then
   txtMontoIngreso.Text = txtValor.Text
Else
   txtMontoIngreso.Text = Format(mMonto * CCur(txtValor) / 100, "Standard")
End If
 
vError:
End Sub


