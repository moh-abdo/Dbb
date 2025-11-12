Attribute VB_Name = "RentalModule"
Option Compare Database
Option Explicit

' RentalModule.bas
' DAO-based helper functions for Access 2016 car rental system

' Check if a car is available for a desired date range
Public Function IsCarAvailable(ByVal lCarID As Long, ByVal dtStart As Date, ByVal dtEnd As Date) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    On Error GoTo ErrHandler
    Set db = CurrentDb()
    strSQL = "SELECT RentalID FROM Rentals WHERE CarID=" & lCarID & _
             " AND ((#" & Format(dtStart, "yyyy\/mm\/dd") & "# <= EndDate) AND (#" & Format(dtEnd, "yyyy\/mm\/dd") & "# >= StartDate))" & _
             " AND Returned = False;"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    IsCarAvailable = rs.EOF
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    IsCarAvailable = False
    Resume Cleanup
End Function

' Calculate rental amount (days * dailyRate + extraFees + lateFees)
Public Function CalculateRentalAmount(ByVal dailyRate As Currency, ByVal dtStart As Date, ByVal dtEnd As Date, _
                                     Optional ByVal extraFees As Currency = 0, Optional ByVal lateFees As Currency = 0) As Currency
    Dim lDays As Long
    lDays = DateDiff("d", dtStart, dtEnd) + 1
    If lDays < 1 Then lDays = 1
    CalculateRentalAmount = (dailyRate * lDays) + extraFees + lateFees
End Function

' Create a rental record and mark car as Rented (returns new RentalID)
Public Function CreateRental(ByVal lCustomerID As Long, ByVal lCarID As Long, ByVal dtStart As Date, ByVal dtEnd As Date, _
                             ByVal dailyRate As Currency, Optional ByVal extraFees As Currency = 0, Optional ByVal sNotes As String = "") As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim lNewID As Long
    On Error GoTo ErrHandler
    Set db = CurrentDb()
    ' Insert rental
    Set rs = db.OpenRecordset("Rentals", dbOpenDynaset)
    rs.AddNew
    rs!CustomerID = lCustomerID
    rs!CarID = lCarID
    rs!StartDate = dtStart
    rs!EndDate = dtEnd
    rs!DailyRate = dailyRate
    rs!ExtraFees = extraFees
    rs!TotalAmount = CalculateRentalAmount(dailyRate, dtStart, dtEnd, extraFees, 0)
    rs!Returned = False
    rs!Notes = sNotes
    rs.Update
    rs.Bookmark = rs.LastModified
    lNewID = rs!RentalID
    rs.Close

    ' Update car status
    db.Execute "UPDATE Cars SET Status='Rented' WHERE CarID=" & lCarID & ";", dbFailOnError

    CreateRental = lNewID
Cleanup:
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    CreateRental = 0
    MsgBox "Error creating rental: " & Err.Description, vbExclamation
    Resume Cleanup
End Function

' Complete a return: set ActualReturnDate, compute late fees, update totals, create payment record(s), set car Available
Public Function CompleteReturn(ByVal lRentalID As Long, ByVal dtActualReturn As Date, Optional ByVal lateFeePerDay As Currency = 20) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim lLateDays As Long
    Dim curLateFees As Currency
    Dim curTotal As Currency
    Dim curDailyRate As Currency
    Dim dtEnd As Date
    Dim lCarID As Long
    On Error GoTo ErrHandler
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM Rentals WHERE RentalID=" & lRentalID, dbOpenDynaset)
    If rs.EOF Then
        MsgBox "Rental not found.", vbExclamation
        CompleteReturn = False
        GoTo Cleanup
    End If
    dtEnd = rs!EndDate
    curDailyRate = rs!DailyRate
    lCarID = rs!CarID
    ' Calculate late days
    If dtActualReturn > dtEnd Then
        lLateDays = DateDiff("d", dtEnd, dtActualReturn)
    Else
        lLateDays = 0
    End If
    curLateFees = lLateDays * lateFeePerDay
    ' Update rental
    rs.Edit
    rs!ActualReturnDate = dtActualReturn
    rs!LateFees = curLateFees
    rs!TotalAmount = CalculateRentalAmount(curDailyRate, rs!StartDate, rs!EndDate, Nz(rs!ExtraFees, 0), curLateFees)
    rs!Returned = True
    rs.Update
    rs.Close

    ' Insert payment record (you can modify to split payments)
    db.Execute "INSERT INTO Payments (RentalID, CustomerID, PaymentDate, Amount, PaymentMethod) " & _
               "SELECT " & lRentalID & ", Rentals.CustomerID, Date(), Rentals.TotalAmount, 'Cash' FROM Rentals WHERE RentalID=" & lRentalID & ";", dbFailOnError

    ' Set car available
    db.Execute "UPDATE Cars SET Status='Available' WHERE CarID=" & lCarID & ";", dbFailOnError

    CompleteReturn = True
Cleanup:
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    MsgBox "Error completing return: " & Err.Description, vbExclamation
    CompleteReturn = False
    Resume Cleanup
End Function

' Quick helper: Nz (for pre-VBA 7; Access has Nz built-in so optional)
Private Function Nz(val, Optional defVal = 0)
    If IsNull(val) Then
        Nz = defVal
    Else
        Nz = val
    End If
End Function
