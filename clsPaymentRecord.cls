Option Explicit

' Class: clsPaymentRecord
' Represents a payment made on a specific day.

Public day As Long
' The day the payment was made.

Public Payment As Double
' The amount of the payment.

Public IsUsed As Boolean
' Indicates if the payment has been used.

Public PaymentID As String
' A unique identifier for the payment.

Private Sub Class_Initialize()
    ' Initializes IsUsed to False when the class is created.
    IsUsed = False
End Sub
