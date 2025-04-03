VERSION 5.00
Begin VB.Form frmGenGeneraAsientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Asientos a Contabilidad"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   7980
   Begin VB.CommandButton cmdParche01 
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "frmGenGeneraAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdParche01_Click()
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = InputBox("Digite clave de Ejecucion ", "Actualizacion")
If Trim(strSQL) <> "trick" Then Exit Sub

'Estable los Saldos a Proveedores, pendientes segun nuevo metodo

strSQL = "select isnull(sum(total),0) as Monto,Cod_Proveedor from cpr_Compras where estado = 'P'" _
       & " and cxp_estado = 'P' and forma_pago = 'CR' group by cod_proveedor"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) + " & rs!Monto _
        & " where cod_proveedor = " & rs!cod_proveedor
 rs.MoveNext
Loop
rs.Close


MsgBox "Parche Ejecutado Satisfactoriamente...", vbInformation

End Sub

