VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARCreditoVigente 
   Caption         =   "ActiveReport1"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ARCreditoVigente.dsx":0000
End
Attribute VB_Name = "ARCreditoVigente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_Format()
    If Not RDODataControl1.Resultset.EOF Then
        With RDODataControl1.Resultset
            fMontoOtorgado.Text = Format(.rdoColumns("ImporteAFinanciar"), "Fixed")
            fCuotas.Text = Format(.rdoColumns("NumCuotas"), "00")
            fVto1raCuota.Text = Format(.rdoColumns("FechaVencimiento1"), "DD/MM/YYYY")
            fAlta.Text = Format(.rdoColumns("FechaCredito"), "DD/MM/YYYY")
        End With
    End If
End Sub

