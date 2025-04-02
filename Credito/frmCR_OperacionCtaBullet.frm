VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmCR_OperacionCtaBullet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definir/Modificar Cuota Balloon de la Operación"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtPlazoRst 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame fraBullet 
      Caption         =   "Parámetros para Definir Cuota Bullet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4080
      TabIndex        =   18
      Top             =   2040
      Width           =   4095
      Begin VB.TextBox txtTasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBulletCta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtBulletAjuste 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         Text            =   "1"
         Top             =   1320
         Width           =   375
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   330
         Left            =   1920
         TabIndex        =   21
         Top             =   1800
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   3387
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplica Cuota Bullet"
               Key             =   "Aplica"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Tasa Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Cuota Bullet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Ajustar Saldos  faltando"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota para Finalizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   22
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   3960
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.TextBox txtBulletActuaAjuste 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtBulletActuaCta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtSaldoBase 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtSaldoReal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtLinea 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.TextBox txtLineaDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_OperacionCtaBullet.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Plazo Restante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   26
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ajusta Saldos Faltando"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cuota Bullet Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo (Base)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo (Real)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblOficina 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oficina...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Label lblOperacion 
      Alignment       =   1  'Right Justify
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8280
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmCR_OperacionCtaBullet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vActiva As Boolean

Private Sub Form_Load()
vModulo = 3

lblOperacion.Caption = Operacion.OperacionConsulta
vActiva = False
End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError
strSQL = "select S.cedula,S.nombre,R.codigo,C.descripcion,R.montoApr,R.Saldo,R.cuota,R.Interesv,R.plazo,R.Estado" _
       & ",R.Dia_Pago,R.Base_Calculo,R.PriDeduc,R.FecUlt,R.int as TasaO,Ofi.descripcion as 'OficinaX'" _
       & ", dbo.fxCrdPlanPagoPlzRestante(R.id_solicitud) as 'PlzRst' ,dbo.fxCrdPlanPagoSldPendientePrg(R.id_solicitud) as 'SaldoPlan'" _
       & ", R.BULLET_IND,isnull(R.BULLET_CTA,0) as 'BulletCta',isnull(R.BULLET_CTA_AJUSTE,1) as 'BulletAjuste'" _
       & ", C.Base_Calculo" _
       & " from Socios S inner join Reg_creditos R on S.cedula = R.cedula" _
       & " inner join catalogo C on R.codigo = C.codigo" _
       & " left join SIF_Oficinas Ofi on R.cod_oficina_r = Ofi.cod_oficina" _
       & " where R.id_solicitud = " & lblOperacion.Caption
Call OpenRecordSet(rs, strSQL)

txtCedula.Text = rs!Cedula
txtNombre.Text = rs!Nombre

txtLinea.Text = rs!Codigo
txtLineaDesc.Text = rs!Descripcion

lblOficina.Caption = rs!OficinaX & ""

txtMonto.Text = Format(rs!montoapr, "Standard")
txtSaldoReal.Text = Format(rs!Saldo, "Standard")

txtTasa.Text = rs!interesv
txtTasa.ToolTipText = "Tasa Original : " & rs!TasaO

txtSaldoBase.Text = Format(rs!SaldoPlan, "Standard")
txtPlazoRst.Text = rs!PlzRst


 If IsNull(rs!Estado) Then
   vActiva = False
   txtSaldoBase.Text = Format(rs!montoapr, "Standard")
 Else
   vActiva = True
 End If

txtBulletActuaAjuste.Text = rs!BulletAjuste
txtBulletActuaCta.Text = Format(rs!BulletCta, "Standard")

'Define Cuota Minima
 txtBulletAjuste.Text = rs!BulletAjuste
 If rs!Base_Calculo = "04" Then
     txtBulletCta.Text = Format(CCur(txtSaldoBase.Text) * 31 * rs!interesv / 36000, "Standard")
 Else
     txtBulletCta.Text = Format(CCur(txtSaldoBase.Text) * 30 * rs!interesv / 36000, "Standard")
 End If
 
 txtBulletCta.Tag = CCur(txtBulletCta.Text)
 txtBulletCta.ToolTipText = "Cta.Min.: " & txtBulletCta.Text

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""

If Not IsNumeric(txtBulletCta.Text) Or Not IsNumeric(txtBulletAjuste.Text) Then
  vMensaje = vMensaje & vbCrLf & " - La cuota bullet o el periodo de ajuste no son válidos..."
End If

If Len(vMensaje) = 0 Then
    'Hace las validaciones correspondientes a los rangos admitidos
    If CLng(txtBulletAjuste.Text) > CLng(txtPlazoRst.Text) Then
     vMensaje = vMensaje & vbCrLf & " - El periodo de Ajuste es mayor que el plazo restante..."
    End If
    If CCur(txtBulletCta.Text) < CCur(txtBulletCta.Tag) Then
     vMensaje = vMensaje & vbCrLf & " - La cuota Bullet es menor que la cuota mínima aplicable..."
    End If
    If CCur(txtBulletCta.Text) > CCur(txtSaldoBase.Text) Then
     vMensaje = vMensaje & vbCrLf & " - La cuota Bullet es mayor que el Saldo Base..."
    End If
End If

If Len(vMensaje) = 0 Then
   fxValida = True
Else
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

End Function

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

If fxValida Then
    strSQL = "update reg_creditos set BULLET_IND = 1, BULLET_CTA = " & CCur(txtBulletCta.Text) _
           & ",BULLET_CTA_AJUSTE = " & txtBulletAjuste.Text _
           & " where id_solicitud = " & lblOperacion.Caption
    Call ConectionExecute(strSQL)
    
     If GLOBALES.SysPlanPagos = 1 And vActiva Then
        'Actualiza Tabla de Pagos
        strSQL = "exec spCrdPlanPagos " & lblOperacion.Caption & ",1"
        Call ConectionExecute(strSQL)
     End If
     
     Call sbBitacoraCredito("22", "De.:" & txtBulletActuaCta.Text & " (Aj." & txtBulletActuaAjuste.Text _
                        & ") A..:" & txtBulletCta.Text & " (Aj." & txtBulletAjuste.Text, "C", lblOperacion.Caption, txtLinea.Text, "")

     MsgBox "Cuota Bullet Establecida o Actualizada satisfactoriamente.!", vbInformation
     Unload Me
End If

End Sub
