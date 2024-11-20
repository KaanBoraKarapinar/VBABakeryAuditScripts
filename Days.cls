Option Explicit

' Class: DAYS
' Manages daily records, payments, and discrepancies across multiple days.

Private mainDict As Scripting.Dictionary
' Dictionary to hold clsDailyRecord objects, keyed by day.

Private centralPaymentManagement As Scripting.Dictionary
' Central management of all payments, keyed by PaymentID.

Private freePaymentMethods As Scripting.Dictionary
' Dictionary of payments that are not yet used, keyed by PaymentID.

Public Sub AddPaymentRecordToDay(day As Long, PaymentRecord As clsPaymentRecord)
    ' Adds a PaymentRecord to a day's record.
    If Not mainDict.Exists(day) Then
        Dim DRtoSave As clsDailyRecord
        Set DRtoSave = New clsDailyRecord
        DRtoSave.day = day
        DRtoSave.addToPaymentRecordsArray PaymentRecord
        mainDict.Add day, DRtoSave
        AddReferenceOfPaymentRecordToTheCentralManagement PaymentRecord
    Else
        mainDict(day).addToPaymentRecordsArray PaymentRecord
        AddReferenceOfPaymentRecordToTheCentralManagement PaymentRecord
    End If
End Sub

Public Sub AddReferenceOfPaymentRecordToTheCentralManagement(PaymentRecord As clsPaymentRecord)
    ' Adds a PaymentRecord to central management and free payments.
    PaymentRecord.PaymentID = "D" & PaymentRecord.day & "A" & PaymentRecord.Payment
    centralPaymentManagement.Add PaymentRecord.PaymentID, PaymentRecord
    freePaymentMethods.Add PaymentRecord.PaymentID, PaymentRecord
End Sub

Public Sub AddSalesRecordToDay(day As Long, SalesRecord As clsSalesRecord)
    ' Adds a SalesRecord to a day's record.
    If Not mainDict.Exists(day) Then
        Dim DRtoSave As clsDailyRecord
        Set DRtoSave = New clsDailyRecord
        DRtoSave.day = day
        Set DRtoSave.SalesRecord = SalesRecord
        mainDict.Add day, DRtoSave
    Else
        If Not mainDict(day).SalesRecord Is Nothing Then
            MsgBox ("Wrong Raw Data: Two Sales Records for same day: " & day)
            Exit Sub
        End If
        Set mainDict(day).SalesRecord = SalesRecord
    End If
End Sub

Property Get Dictionary() As Scripting.Dictionary
    ' Getter for the main dictionary of days.
    Set Dictionary = mainDict
End Property

Public Sub CheckAllDiscrepancy()
    ' Checks discrepancies for all days.
    Dim current As Variant
    Dim currentDay As clsDailyRecord

    For Each current In mainDict.Keys
        Set currentDay = mainDict(current)

        If current = 1 Then
            currentDay.ApplyAllPaymentsOfDayAsEffective
            currentDay.CalculateDiscrepancyWithCurrentEffectivePayments
            ForceRefreshFreePaymentRecords
        Else
            If Not currentDay.SalesRecord Is Nothing Then
                currentDay.ApplyAllRemainingPaymentsOfDayAsEffective
                ForceRefreshFreePaymentRecords
            End If
        End If

        currentDay.CalculateDiscrepancyWithCurrentEffectivePayments

        If currentDay.IsDiscrepant Then
            Dim discrepancyFiller As clsPaymentRecord
            Set discrepancyFiller = FindPaymentWithPaymentAmount(currentDay.Discrepancy)

            If Not discrepancyFiller Is Nothing Then
                currentDay.AddEffectivePayment discrepancyFiller
                currentDay.CalculateDiscrepancyWithCurrentEffectivePayments
                ForceRefreshFreePaymentRecords
            End If
        End If
    Next
End Sub

Public Sub applyAllRemaining()
    ' Applies all remaining payments as effective.
    Dim current As Variant
    Dim currentDay As clsDailyRecord

    For Each current In mainDict.Keys
        Set currentDay = mainDict(current)
        currentDay.ApplyAllRemainingPaymentsOfDayAsEffective
    Next
End Sub

Public Sub ForceRefreshFreePaymentRecords()
    ' Refreshes the list of free payment methods.
    Set freePaymentMethods = New Scripting.Dictionary
    Dim currentID As Variant

    For Each currentID In centralPaymentManagement.Keys
        If centralPaymentManagement(currentID).IsUsed = False Then
            freePaymentMethods.Add currentID, centralPaymentManagement(currentID)
        End If
    Next
End Sub

Public Function FindPaymentWithPaymentAmount(Amount As Long) As clsPaymentRecord
    ' Finds a free payment record with a specific payment amount.
    On Error Resume Next
    Dim currentID As Variant
    Dim returnRecord As clsPaymentRecord

    For Each currentID In freePaymentMethods.Keys
        If freePaymentMethods(currentID).IsUsed = False And freePaymentMethods(currentID).Payment = Amount Then
            Set returnRecord = freePaymentMethods(currentID)
            Exit For
        End If
    Next

    Set FindPaymentWithPaymentAmount = returnRecord
End Function

Public Sub BasicCheckAllDiscrepancy()
    ' Performs basic discrepancy checks for all days.
    Dim currentDay As Variant

    For Each currentDay In mainDict.Keys
        mainDict(currentDay).BasicCheckDiscrepancy
    Next
End Sub

Public Sub checkAllDiscrepancyAllowOnlyNextDay()
    ' Checks discrepancies, allowing payments from the next day.
    Dim currentDay As Variant

    For Each currentDay In mainDict.Keys
        On Error Resume Next
        Dim previousDay As Long
        Dim effectivePaymentToday As Double
        Dim calculativePaymentToday As Double

        calculativePaymentToday = mainDict(currentDay).TotalPaymentAmount
        previousDay = currentDay - 1

        If mainDict.Exists(previousDay) Then
            If mainDict(previousDay).IsDiscrepant Then
                Dim discrepancyPreviousDay As Long
                discrepancyPreviousDay = mainDict(previousDay).Discrepancy

                If calculativePaymentToday >= discrepancyPreviousDay Then
                    mainDict(previousDay).Discrepancy = 0
                    mainDict(previousDay).IsDiscrepant = False
                    effectivePaymentToday = calculativePaymentToday - discrepancyPreviousDay
                Else
                    mainDict(previousDay).Discrepancy = discrepancyPreviousDay - calculativePaymentToday
                    effectivePaymentToday = 0
                End If
            Else
                effectivePaymentToday = calculativePaymentToday
            End If
        Else
            effectivePaymentToday = calculativePaymentToday
        End If

        mainDict(currentDay).ComplexCheckDiscrepancy effectivePaymentToday
    Next
End Sub

Function getAnArrayOfSortedKeys() As Long()
    ' Returns an array of sorted keys (days).
    Dim k As Variant
    Dim i0 As Long
    i0 = 0
    Dim SortedArray() As Long

    ReDim SortedArray(0 To mainDict.Count - 1)

    For Each k In mainDict.Keys()
        SortedArray(i0) = k
        i0 = i0 + 1
    Next

    quicksort SortedArray, 0, UBound(SortedArray)
    getAnArrayOfSortedKeys = SortedArray
End Function

Sub quicksort(a() As Long, left As Long, right As Long)
    ' QuickSort algorithm.
    Dim P As Long

    If left < right Then
        P = partition(a, left, right)
        quicksort a, left, P
        quicksort a, P + 1, right
    End If
End Sub

Function partition(a() As Long, left As Long, right As Long) As Long
    ' Partition function for QuickSort.
    Dim Pivot As Long
    Pivot = a(left)
    Dim pl As Long
    Dim pr As Long
    pl = left
    pr = right

    Do While pl < pr
        Do While a(pl) <= Pivot And pl < right
            pl = pl + 1
        Loop

        Do While a(pr) > Pivot And pr > left
            pr = pr - 1
        Loop

        If pl < pr Then
            swap a, pl, pr
        End If
    Loop

    swap a, left, pr
    partition = pr
End Function

Sub swap(a() As Long, i As Long, j As Long)
    ' Swaps two elements in an array.
    Dim temp As Long
    temp = a(i)
    a(i) = a(j)
    a(j) = temp
End Sub

Sub WriteData()
    ' Writes data to the worksheet starting at cell A10.
    On Error Resume Next
    Dim startIndex As Range
    Set startIndex = Range("A10")
    Dim ColumnPointer As Long
    ColumnPointer = 1
    Dim sortedKeys As Variant
    sortedKeys = Me.getAnArrayOfSortedKeys
    Dim Key As Variant

    Range("A10", "D100").Value = ""

    For Each Key In sortedKeys
        startIndex.Cells(ColumnPointer, 1).Value = mainDict(Key).day
        startIndex.Cells(ColumnPointer, 2).Value = mainDict(Key).SalesRecord.SalesAmount
        startIndex.Cells(ColumnPointer, 3).Value = mainDict(Key).TotalPaymentAmount
        startIndex.Cells(ColumnPointer, 4).Value = mainDict(Key).Discrepancy
        ColumnPointer = ColumnPointer + 1
    Next Key
End Sub

Sub WriteAuditResult()
    ' Writes the audit results to cell B4.
    Dim Index As Range
    Set Index = Range("B4")
    Dim Result As String
    Dim sortedKeys As Variant
    sortedKeys = Me.getAnArrayOfSortedKeys
    Dim Key As Variant

    For Each Key In sortedKeys
        If mainDict(Key).IsDiscrepant Then
            Result = Result & " " & mainDict(Key).day
        End If
    Next Key

    Index.Value = Result
End Sub

Sub ReverseArray(a() As Long)
    ' Reverses the elements of an array.
    Dim left As Long
    Dim right As Long
    left = LBound(a)
    right = UBound(a)

    Do While left < right
        swap a, left, right
        left = left + 1
        right = right - 1
    Loop
End Sub

Private Sub Class_Initialize()
    ' Initializes the dictionaries.
    Set mainDict = New Scripting.Dictionary
    Set centralPaymentManagement = New Scripting.Dictionary
    Set freePaymentMethods = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    ' Cleanup code if necessary.
End Sub
