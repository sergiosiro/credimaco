VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARCreditosNuevo 
   Caption         =   "Imprimir creditos"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "ARCreditosNuevo.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARCreditosNuevo.dsx":0442
End
Attribute VB_Name = "ARCreditosNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
On Error GoTo merror


With RDODataControl1.Resultset
     If .rdoColumns("aplicarcuotacomodin") Then
        FieldCuotaComodin.Text = "SI"
     Else
        FieldCuotaComodin.Text = "NO"
     End If
     
End With


Exit Sub
merror:
tratarerrores "Error en reporte ARCreditosLista"
End Sub


