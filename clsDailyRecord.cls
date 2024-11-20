Option Explicit

' Class: clsDailyRecord
' Represents a daily record containing payments and sales.
' Calculates discrepancies between sales and payments.

Public day As Long
' The day number (e.g., 1, 2, 3, etc.)

Private PaymentRecord() As clsPaymentRecord
' Array of PaymentRecords for the day.

Private EffectivePaymentRecord() As clsPaymentRecord
' Array of Effective PaymentRecords used for discrepancy calculation.

Public TotalNumberOfEffectivePayments As Long
' Total number of effective payments for the day.

Public TotalEffectivePaymentAmount As Long
' Total amount of effective payments for the day.

Public SalesRecord As clsSalesRecord
' SalesRecord for the day.

Public Discrepancy As Double
' The discrepancy amount for the day.

Public IsDiscrepant As Boolean
' Indicates if there is a discrepancy on this day.

Public ThereIsSavedPayment As Boolean
' Indicates if there are any saved payments for the day.

Public TotalPaymentAmount As Double
' Total amount of payments for the day.

Private TotalNumberOfPayment As Long
' Total number of payments for the day.

Sub ToString()
    ' Displays information about the day's record.
    On Error Resume Next
    MsgBox "In day: " & day & " payment: " & PaymentRecord.Payment & " sales: " & SalesRecord.SalesAmount & _
           " isDiscrepant and Amount: " & IsDiscrepant & " " & Discrepancy
End Sub

Public Property Get getReadOnlyPaymentRecords() As Long
    ' Getter for PaymentRecord array (read-only).
    paymentrecords = PaymentRecord
End Property

Public Sub addToPaymentRecordsArray(value As clsPaymentRecord)
    ' Adds a PaymentRecord to the PaymentRecord array.
    ThereIsSavedPayment = True
    TotalNumberOfPayment = TotalNumberOfPayment + 1
    ReDim Preserve PaymentRecord(0 To TotalNumberOfPayment - 1)
    Set PaymentRecord(TotalNumberOfPayment - 1) = value
    TotalPaymentAmount = TotalPaymentAmount + value.Payment
End Sub

Public Sub AddEffectivePayment(effectivePayment As clsPaymentRecord)
    ' Adds a PaymentRecord to the EffectivePaymentRecord array if it hasn't been used yet.
    If effectivePayment.IsUsed = False Then
        TotalNumberOfEffectivePayments = TotalNumberOfEffectivePayments + 1
        Dim lastLegalIndex As Long
        lastLegalIndex = TotalNumberOfEffectivePayments - 1
        ReDim Preserve EffectivePaymentRecord(0 To lastLegalIndex)
        Set EffectivePaymentRecord(lastLegalIndex) = effectivePayment
        TotalEffectivePaymentAmount = TotalEffectivePaymentAmount + effectivePayment.Payment
        effectivePayment.IsUsed = True
    Else
        ' MsgBox "Used payments cannot be effective"
    End If
End Sub

Public Sub ApplyAllPaymentsOfDayAsEffective()
    ' Makes all payments of the day effective.
    Dim current As Variant
    Dim tempRecord As clsPaymentRecord
    For Each current In PaymentRecord
        Set tempRecord = current
        AddEffectivePayment tempRecord
    Next
End Sub

Public Sub ApplyAllRemainingPaymentsOfDayAsEffective()
    ' Makes all unused payments of the day effective.
    Dim current As Variant
    Dim tempRecord As clsPaymentRecord
    For Each current In PaymentRecord
        If current.IsUsed = False Then
            Set tempRecord = current
            AddEffectivePayment tempRecord
        End If
    Next
End Sub

Public Sub makeBiggestPaymentEffective()
    ' Makes the biggest unused payment effective.
    Dim tempIndex As Long
    tempIndex = IndexOfBiggestNonUsedPayment
    If tempIndex <> -1 Then
        AddEffectivePayment PaymentRecord(tempIndex)
    End If
End Sub

Public Sub CalculateDiscrepancyWithCurrentEffectivePayments()
    ' Calculates discrepancy based on current effective payments and sales record.
    If SalesRecord Is Nothing Then
        IsDiscrepant = True
        Discrepancy = -TotalEffectivePaymentAmount
        If TotalEffectivePaymentAmount = 0 Then
            IsDiscrepant = False
            Discrepancy = -TotalEffectivePaymentAmount
        End If
        Exit Sub
    ElseIf TotalEffectivePaymentAmount = 0 Then
        IsDiscrepant = True
        Discrepancy = SalesRecord.SalesAmount
        Exit Sub
    End If

    If SalesRecord.SalesAmount = TotalEffectivePaymentAmount Then
        Discrepancy = 0
        IsDiscrepant = False
    Else
        IsDiscrepant = True
        Discrepancy = SalesRecord.SalesAmount - TotalEffectivePaymentAmount
    End If
End Sub

Public Function IndexOfBiggestNonUsedPayment() As Long
    ' Returns the index of the biggest non-used payment.
    Dim i As Long
    Dim maxIndex As Long
    Dim maxValue As Double
    maxIndex = -1
    maxValue = -1

    For i = 0 To UBound(PaymentRecord)
        If Not PaymentRecord(i).IsUsed And PaymentRecord(i).Payment > maxValue Then
            maxValue = PaymentRecord(i).Payment
            maxIndex = i
        End If
    Next i

    IndexOfBiggestNonUsedPayment = maxIndex
End Function

Public Function getNumberOfNotUsedPayments() As Long
    ' Returns the number of payments that have not been used.
    Dim count As Long
    count = 0
    For Each Payment In PaymentRecord
        If Payment.IsUsed = False Then
            count = count + 1
        End If
    Next
    getNumberOfNotUsedPayments = count
End Function

Sub BasicCheckDiscrepancy()
    ' Performs a basic discrepancy check between sales and total payments.
    If SalesRecord Is Nothing And Not ThereIsSavedPayment Then
        IsDiscrepant = False
        Exit Sub
    ElseIf SalesRecord Is Nothing Then
        IsDiscrepant = True
        Discrepancy = -TotalPaymentAmount
        Exit Sub
    ElseIf Not ThereIsSavedPayment Then
        IsDiscrepant = True
        Discrepancy = SalesRecord.SalesAmount
        Exit Sub
    End If

    If SalesRecord.SalesAmount = TotalPaymentAmount Then
        IsDiscrepant = False
    Else
        IsDiscrepant = True
        Discrepancy = SalesRecord.SalesAmount - TotalPaymentAmount
    End If
End Sub

Sub ComplexCheckDiscrepancy(effectivePayment As Double)
    ' Performs a discrepancy check using effective payments.
    If SalesRecord Is Nothing And Not ThereIsSavedPayment Then
        IsDiscrepant = False
        Exit Sub
    ElseIf SalesRecord Is Nothing Then
        IsDiscrepant = True
        Discrepancy = -effectivePayment
        Exit Sub
    ElseIf Not ThereIsSavedPayment Then
        IsDiscrepant = True
        Discrepancy = SalesRecord.SalesAmount
        Exit Sub
    End If

    If SalesRecord.SalesAmount = effectivePayment Then
        IsDiscrepant = False
    Else
        IsDiscrepant = True
        Discrepancy = SalesRecord.SalesAmount - effectivePayment
    End If
End Sub

Sub CheckDiscrepancy()
    ' Empty Sub (possibly for future use).
End Sub

Private Sub Class_Initialize()
    ' Initializes arrays and counters.
    ReDim PaymentRecord(0 To 0)
    TotalNumberOfPayment = 0
    TotalPaymentAmount = 0
    TotalEffectivePaymentAmount = 0
    TotalNumberOfEffectivePayments = 0
End Sub
