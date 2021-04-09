Attribute VB_Name = "Module1"
Option Explicit

' Function to judge Sunday and special holiday.
Public Function Holiday(ByVal ReqDate As Date) As Boolean
Dim SPHolobj As Holidaycls
Dim WD As Integer
Dim HolName As String

    WD = Weekday(ReqDate)
    
    If WD = vbSunday Then
        Holiday = True
        Exit Function
    End If
    
    Set SPHolobj = New Holidaycls
    
    If SPHolobj.Holiday(ReqDate, HolName) > 0 Then
        Holiday = True
        Exit Function
    End If
    
    Holiday = False
    
End Function

Public Function HolidayName(ByVal ReqDate As Date) As String
Dim SPHolobj As Holidaycls
Dim HolName As String
    
    Set SPHolobj = New Holidaycls
    
    If SPHolobj.Holiday(ReqDate, HolName) > 0 Then
        HolidayName = HolName
        Exit Function
    End If
    
    HolidayName = ""
    
End Function
