VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AcuerdoMutuo 
   Caption         =   "Mutuo Acuerdo"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "AcuerdoMutuo.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "AcuerdoMutuo.dsx":0442
End
Attribute VB_Name = "AcuerdoMutuo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
With RDODataControl1.Resultset
     TxtNumCuota.Text = CStr(.rdoColumns("numcuota")) & "/" & CStr(.rdoColumns("numcuotas"))
End With
End Sub

Private Sub PageFooter_Format()
Field2.Text = Me.pageNumber
End Sub

