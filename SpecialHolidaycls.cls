VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecialHolidaycls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class of Special Hohliday.
' Refferd to http://www.asahi-net.or.jp/~ef2o-inue/vba_o/sub05_800_052.html.
Option Explicit

Private mStartYear As Integer ' The start year when the holiday is applied.
Private mEndYear As Integer ' The last year when the holiday is applied.
Private mMonth As Integer   ' The month of the holiday.
Private mDayofMonth As Integer  ' The day of the month of the holiday.
Private mAltShift As Integer    ' The shift number of day for alternate holiday.
Private mNumofWeek As Integer   ' The Number of the week when the holiday is (for so as call a Happy Monday).
Private mIsVernal_E As Boolean  ' The flag for Vernal Equinox.
Private mIsAutumn_E As Boolean  ' The flag for Autumn Equinox.
Private mName As String         ' The name of the holiday.
Private mPrevHolidayObj As SpecialHolidaycls ' The previous object for the SpecialHolidaycls.

' Subroutine to initiate the class.
' The month of the holiday is specified on the Argument of ReqMonth(zero is default value).
' The day of the month of the holiday is specified on the Argumnet of ReqDayofMonth(zero is default value).
' The shift number of day for alternate holiday is specified on the Argumnet of ReqAltShift(one is default value).
' The Number of the week when the holiday is (for so as call Happy Monday) is specified on the Argumnet of ReqNumofWeek
' (zero is default value).
' The Vernal Equinox is specified on ReqIsVernal_E(false is default value).
' The Autumn Equinox is specified on ReqIsAutumn_E(false is default value).
Public Sub Init(Optional ByVal ReqStartYear As Integer = 1948, Optional ByVal ReqEndYear As Integer = 2150, _
                    Optional ByVal ReqMonth As Integer = 0, Optional ByVal ReqDayofMonth As Integer = 0, _
                    Optional ByVal ReqAltShift As Integer = 1, _
                    Optional ByVal ReqNumofWeek As Integer = 0, _
                    Optional ByVal ReqIsVernal_E As Boolean = False, Optional ByVal ReqIsAutumn_E As Boolean = False, _
                    Optional ByVal ReqPrevSHObj As SpecialHolidaycls = Nothing, Optional ByVal ReqName As String = "国民の休日")
Dim CheckDate As Date
Dim CheckDay As Integer

    mDayofMonth = 0
    mAltShift = 0
    mNumofWeek = 0
    
    mIsVernal_E = False
    mIsAutumn_E = False
    
    ' Value check for argument of ReqStartYear.
    If (ReqStartYear < 1948 Or ReqStartYear > 2150) Then
        Err.Raise Number:=888, Source:="祝日発生ルーチン", Description:="エラー：対応する年は1948年から2150年までです。"
        Exit Sub
    Else
        mStartYear = ReqStartYear
    End If
    
    ' Value check for argument of ReqEndYear.
    If (ReqEndYear < 1948 Or ReqEndYear > 2150) Then
        Err.Raise Number:=889, Source:="祝日発生ルーチン", Description:="エラー：対応する年は1948年から2150年までです。"
        Exit Sub
    Else
        mEndYear = ReqEndYear
    End If
    
    ' Error of value of argument of Month.
    If (ReqMonth < 1 Or ReqMonth > 12) And (ReqIsVernal_E = False And ReqIsAutumn_E = False) Then
        Err.Raise Number:=890, Source:="祝日発生ルーチン", Description:="エラー：月の数字は1～12までを指定してください。"
        Exit Sub
    Else
        mMonth = ReqMonth
    End If
    
    ' When any kinds of the holiday is not specified.
    If ReqDayofMonth < 1 And ReqNumofWeek < 1 And ReqIsVernal_E = False And ReqIsAutumn_E = False Then
        Err.Raise Number:=891, Source:="祝日発生ルーチン", Description:="エラー：祝日の種類を特定できません。"
        Exit Sub
    End If
    
    ' When the both kinds of the holiday is specified.
    If ReqDayofMonth > 0 And ReqNumofWeek > 0 Then
        Err.Raise Number:=892, Source:="祝日発生ルーチン", Description:="エラー：祝日の種類を特定できません。"
        Exit Sub
    End If
    
    ' When the both Vernal nad Autumn Equinox is specified.
    If ReqIsVernal_E = True And ReqIsAutumn_E = True Then
        Err.Raise Number:=893, Source:="祝日発生ルーチン", Description:="エラー：春分の日と秋分の日が同時に指定されています。"
        Exit Sub
    End If
    
    If ReqDayofMonth > 0 Then
        CheckDate = DateSerial(2008, ReqMonth, ReqDayofMonth)
        CheckDay = Day(CheckDate)
        If Not ReqDayofMonth = CheckDay Then
            Err.Raise Number:=894, Source:="祝日発生ルーチン", _
                Description:="エラー：" & Format(ReqMonth, "0") & "月" & Format(ReqDayofMonth) & "日は存在しません。"
            Exit Sub
        End If
    End If
    
    If ReqDayofMonth < 1 And ReqNumofWeek < 1 And ReqNumofWeek > 4 Then
        Err.Raise Number:=895, Source:="祝日発生ルーチン", _
            Description:="エラー：週の数は1から4までにしてください。"
        Exit Sub
    End If
    
    mDayofMonth = ReqDayofMonth
    mAltShift = ReqAltShift
    mNumofWeek = ReqNumofWeek
    
    mIsVernal_E = ReqIsVernal_E
    mIsAutumn_E = ReqIsAutumn_E
    
    mName = ReqName
    Set mPrevHolidayObj = ReqPrevSHObj
End Sub

' The function distinguish whether the date in the argument is the holiday.
Public Function Check(ByVal ReqDate As Date, ByRef OutDayName As String) As Integer
Dim RDDate As Date       ' The value which ReqDate is rounded down.
Dim YearVal As Integer   ' Year value of ReqDate.
Dim MonthVal As Integer  ' Month value of Reqdate
Dim DateVal As Integer   ' Day value of ReqDate.
Dim dumDate As Date      ' Date type variable for work.
Dim Befdate As Date      ' Date variable one day before ReqDate.
Dim AftDate As Date      ' Date variable one day after ReqDate.
Dim BefdayName As String ' Day name one day before ReqDate.
Dim AftdayName As String ' Day name one day after ReqDate.

    RDDate = Int(ReqDate)
    
    YearVal = Year(RDDate)
    MonthVal = Month(RDDate)
    DateVal = Day(RDDate)
    
    ' Year Check.
    If YearVal < mStartYear Or YearVal > mEndYear Then
        Check = 0
        OutDayName = ""
        Exit Function
    End If
    
    ' Check of the date specified holiday and its alternate holiday.
    If mDayofMonth > 0 Then
        If MonthVal = mMonth And DateVal = mDayofMonth Then
            Check = 1
            OutDayName = mName
            Exit Function
        End If
        If AltHoliday(ReqDate) = True Then
            Check = 2
            OutDayName = "振替休日"
            Exit Function
        End If
    End If
    
    ' Check of the Happy Monday.
    If mNumofWeek > 0 Then
        If Abs(Happy_Mon(YearVal) - RDDate) < 0.1 Then
            Check = 1
            OutDayName = mName
            Exit Function
        End If
    End If
    
    'Check the Day of Vernal Equinox.
    If mIsVernal_E = True Then
        dumDate = Vernal_Equinox(YearVal)
        If Abs(RDDate - dumDate) < 0.1 Then
            Check = 1
            OutDayName = mName
            Exit Function
        End If
        If Weekday(dumDate) = vbSunday Then
            If Abs(dumDate + 1# - RDDate) < 0.1 And Year(dumDate) > 1972 Then
                Check = 2
                OutDayName = "振替休日"
                Exit Function
            End If
        End If
    End If
    
    'Check of the Autumn Equinox.
    If mIsAutumn_E = True Then
        dumDate = Autumn_Equinox(YearVal)
        If Abs(RDDate - dumDate) < 0.1 Then
            Check = 1
            OutDayName = mName
            Exit Function
        End If
        If Weekday(dumDate) = vbSunday Then
            If Abs(dumDate + 1# - RDDate) < 0.1 And Year(dumDate) > 1972 Then
                Check = 2
                OutDayName = "振替休日"
                Exit Function
            End If
        End If
    End If
    
    ' The day between two holidays have been also holiday since 1985.
    If Not mPrevHolidayObj Is Nothing And YearVal > 1985 Then
        Befdate = Int(ReqDate - 1#)
        AftDate = Int(ReqDate + 1#)
        BefdayName = ""
        AftdayName = ""
        If mPrevHolidayObj.Check(Befdate, BefdayName) > 0 And Me.Check(AftDate, AftdayName) > 0 And Weekday(ReqDate, vbSunday) > 1 Then
            Check = 3
            OutDayName = "国民の休日"
            Exit Function
        End If
    End If
    
    Check = 0
    OutDayName = ""
End Function


' Private functions.

' Function to calculate the Holiday of the Happy Monday.
Private Function Happy_Mon(ByVal ReqYear As Integer) As Date
Dim Ds As Date
Dim FirstWD As Integer
Dim WeekCnt As Integer

    ' The first day of the month.
    Ds = DateSerial(ReqYear, mMonth, 1)
    
    ' The week day of the first day on the month.
    FirstWD = Weekday(Ds, vbSunday)
    
    ' Mondays are after the first week when the first day of month is not Sunday or Monday.
    If FirstWD > vbMonday Then
        WeekCnt = mNumofWeek + 1
    Else
        WeekCnt = mNumofWeek
    End If
    
    ' Caliculate the Happy Monday.
    Happy_Mon = Ds - FirstWD + (WeekCnt - 1) * 7 + vbMonday
End Function

' Function to calculate the Alternate holiday.
Private Function AltHoliday(ByVal ReqDate As Date) As Boolean
Dim dumDate As Date
Dim dumWeek As Integer
    
    ' Add for closed days in end and start of year.
    If mAltShift < 1 Then
        AltHoliday = False
        Exit Function
    End If
    
    dumDate = DateSerial(Year(ReqDate), mMonth, mDayofMonth)
    dumWeek = Weekday(dumDate)
    If dumWeek = vbSunday Then
        ' Alternate holidays have founded since Apr. 12, 1973.
        If Month(ReqDate) = mMonth And Day(ReqDate) = mDayofMonth + mAltShift And dumDate > DateSerial(1973, 4, 11) Then
            AltHoliday = True
            Exit Function
        End If
    End If
    AltHoliday = False
End Function

' Function to output the day of Vernal Equinox.
Private Function Vernal_Equinox(ByVal ReqYear As Integer) As Date
Dim DayNum As Integer
Dim intY2 As Integer

    ' Not covered before 1948 And after 2150.
    intY2 = ReqYear - 1980
    Select Case ReqYear
        Case Is <= 1979
            DayNum = Int(20.8357 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Is <= 2099
            DayNum = Int(20.8431 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Else
            DayNum = Int(21.851 + (0.242194 * intY2) - Int(intY2 / 4))
    End Select
    Vernal_Equinox = DateSerial(ReqYear, 3, DayNum)
End Function

'Function to output the day of Autumn_Equinox.
Private Function Autumn_Equinox(ByVal ReqYear As Integer) As Date
Dim DayNum As Integer
Dim intY2 As Integer

    ' Not covered before 1948 And after 2150.
    intY2 = ReqYear - 1980
    Select Case ReqYear
        Case Is <= 1979
            DayNum = Int(23.2588 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Is <= 2099
            DayNum = Int(23.2488 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Else
            DayNum = Int(24.2488 + (0.242194 * intY2) - Int(intY2 / 4))
    End Select
    Autumn_Equinox = DateSerial(ReqYear, 9, DayNum)
End Function
