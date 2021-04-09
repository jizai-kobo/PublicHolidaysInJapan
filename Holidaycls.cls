VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Holidaycls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Holidays Collection.
Private mHolidaysCollection As Collection

' The function to make collection
' The function to judge the holidays.
' The argument ReqDate is date value to check whether it is holiday, aternete holiday or other day.
' When the argument ReqDate is holiday, function return the value of 1.
' When the argument ReqDate is alternate holiday, function return the value of 2.
' When the argument ReqDate is additional holiday between two holidays, function return the value of 3.
' When the argument ReqDate is the other day, function return the value of 0.
Public Function Holiday(ByVal ReqDate As Date, ByRef OutHolName As String) As Integer
Dim Check As Integer
Dim k As Integer
Dim SHObj As SpecialHolidaycls
Dim HolName As String
Dim MonthStr As String
Dim NextDayMonthStr As String

    k = 0
    HolName = ""
    MonthStr = Format(ReqDate, "mmm")
    Do
        k = k + 1
        Set SHObj = Nothing
        On Error Resume Next
        Set SHObj = mHolidaysCollection.Item(MonthStr & Format(k, "00"))
        On Error GoTo 0
        If SHObj Is Nothing Then
            NextDayMonthStr = Format(ReqDate + 1, "mmm")
            If StrComp(MonthStr, NextDayMonthStr) = 0 Then
                Exit Do
            Else
                MonthStr = NextDayMonthStr
                k = 0
            End If
        Else
            Check = SHObj.Check(ReqDate, HolName)
            If Check > 0 Then
                Holiday = Check
                OutHolName = HolName
                Exit Function
            End If
        End If
    Loop
    
    Holiday = 0
End Function

' Initialize the holidays.
Private Sub Class_Initialize()
' Object instances for each holidays.(After 1948)

' The New year day(Jan. 1 1949-).
Dim NewYear As SpecialHolidaycls

' Coming of age day
Dim ComingofAgeJan15 As SpecialHolidaycls            ' Jan. 15 1949-1997.
Dim ComingofAgeJan2ndMon As SpecialHolidaycls        ' Second Monday in Jan. 2000-.

' The national foundation day
Dim NationalFoundationFeb11 As SpecialHolidaycls     ' Feb. 11 1967-.

' The vernal equinox
Dim VernalEquinox As SpecialHolidaycls               ' Mar. 1949-.

' The day for green.
Dim GreenApr29 As SpecialHolidaycls                  ' Apr. 29 1989-2004.
Dim GreenMay4 As SpecialHolidaycls                   ' May 4 2005-.

' The day of Showa.
Dim DayofShowa As SpecialHolidaycls                  ' Apr. 29 2005-.

' The day of constitution.
Dim Constitution1949 As SpecialHolidaycls            ' May 3 1949-2004.
Dim Constitution2007 As SpecialHolidaycls            ' May 3 2005-.

' The day of children.
Dim Children1949 As SpecialHolidaycls                ' May 5 1949-1987.
Dim Children1988 As SpecialHolidaycls                ' May 5 1988-2006.
Dim Children2007 As SpecialHolidaycls                ' May 5 2007-.

' The day of sea.
Dim DayofSea1996 As SpecialHolidaycls                ' Jul. 20 1996-2002.
Dim DayofSea2003 As SpecialHolidaycls                ' Third Monday in Jul. 2003-2019.
Dim DayofSea2020 As SpecialHolidaycls                ' Jul. 23 2020.
Dim DayofSea2021 As SpecialHolidaycls                ' Jul. 22 2021.
Dim DayofSea2022 As SpecialHolidaycls                ' Third Monday in Jul. 2022-.

' The day of mountains.
Dim DayofMount2016 As SpecialHolidaycls              ' Aug. 11 2016-2019.
Dim DayofMount2020 As SpecialHolidaycls              ' Aug. 10 2020.
Dim DayofMount2021 As SpecialHolidaycls              ' Aug. 8 2021.
Dim DayofMount2022 As SpecialHolidaycls              ' Aug. 11 2022-.

' The day of seniors.
Dim DayofSeniorSep15 As SpecialHolidaycls             ' Sep. 15 1966-.
Dim DayofSeniorSepTrdMon As SpecialHolidaycls         ' Third monday of Sep. 15 2000-.

' The autumn equinox
Dim AutumnEquinox1948 As SpecialHolidaycls             ' Sep. 1948-1987.
Dim AutumnEquinox1988 As SpecialHolidaycls             ' Sep. 1988-2002.
Dim AutumnEquinox2003 As SpecialHolidaycls             ' Sep. 2003-.

' The day of sports.
Dim DayofSportsOct10 As SpecialHolidaycls              ' Oct. 10 1966-1999
Dim DayofSports2000 As SpecialHolidaycls               ' Second Monday in Oct. 2000-2019.
Dim DayofSports2020 As SpecialHolidaycls               ' Jul. 24 2020.
Dim DayofSports2021 As SpecialHolidaycls               ' Jul. 23 2020.
Dim DayofSportsOct2ndMon As SpecialHolidaycls           ' Second Monday in Oct. 2022.

' The day of culture.
Dim Culture As SpecialHolidaycls                       ' Nov. 3 1948-.

' The appreciation day for work.
Dim Work As SpecialHolidaycls                          ' Nov. 23 1948-.

' The birthday of the Emperor.
Dim EmperorApr29 As SpecialHolidaycls                  ' Apr. 29 1949-1988.
Dim EmperorDec23 As SpecialHolidaycls                  ' Dec. 23 1989-2018.
Dim EmperorFeb23 As SpecialHolidaycls                   ' Feb. 23 2020.

' The marriage of the prince in 1959.
Dim MarriageofPrince1959 As SpecialHolidaycls          ' Apr. 10 1959.

' The funeral of the emperor in 1989.
Dim FuneralofEmperor1989 As SpecialHolidaycls          ' Feb. 24 1989.

' The enthronemento of the emperor in 1990.
Dim EnthronementofEmperor1990 As SpecialHolidaycls     ' Nov. 12 1990.

' The marriage of prince in 1993.
Dim MarriageofPrince1993 As SpecialHolidaycls          ' Jun. 9 1993.

' The enthronemento of emperor in 2019.
Dim EnthronementofEmperor2019 As SpecialHolidaycls     ' May 1 2019.

' The enthronemento ceremony in 2019.
Dim EnthronementCeremony2019 As SpecialHolidaycls      ' Oct 22 2019.

' Key Generator.
Dim KeyGenerator As KeyGencls

    Set mHolidaysCollection = New Collection
    Set KeyGenerator = New KeyGencls
    
    ' New Year Day (Jan. 1 1949-).
    Set NewYear = New SpecialHolidaycls
    NewYear.Init ReqStartYear:=1949, ReqMonth:=1, ReqDayofMonth:=1, ReqName:="å≥ì˙"
    mHolidaysCollection.Add Item:=NewYear, Key:=KeyGenerator.Key("Jan")
    
    ' The Coming of Age.
    ' Jan. 15 1949-1999.
    Set ComingofAgeJan15 = New SpecialHolidaycls
    ComingofAgeJan15.Init ReqStartYear:=1949, ReqEndYear:=1999, ReqMonth:=1, ReqDayofMonth:=15, ReqName:="ê¨êlÇÃì˙"
    mHolidaysCollection.Add Item:=ComingofAgeJan15, Key:=KeyGenerator.Key("Jan")
    
    ' Second Monday in Jan. 2000-.
    Set ComingofAgeJan2ndMon = New SpecialHolidaycls
    ComingofAgeJan2ndMon.Init ReqStartYear:=2000, ReqMonth:=1, ReqNumofWeek:=2, ReqName:="ê¨êlÇÃì˙"
    mHolidaysCollection.Add Item:=ComingofAgeJan2ndMon, Key:=KeyGenerator.Key("Jan")
    
    ' National Foundation Day.
    ' Feb. 11 1967-
    Set NationalFoundationFeb11 = New SpecialHolidaycls
    NationalFoundationFeb11.Init ReqStartYear:=1967, ReqMonth:=2, ReqDayofMonth:=11, ReqName:="åöçëãLîOÇÃì˙"
    mHolidaysCollection.Add Item:=NationalFoundationFeb11, Key:=KeyGenerator.Key("Feb")
    
    ' The Vernal Equinox(Mar. 1949-).
    Set VernalEquinox = New SpecialHolidaycls
    VernalEquinox.Init ReqStartYear:=1949, ReqIsVernal_E:=True, ReqName:="ètï™ÇÃì˙"
    mHolidaysCollection.Add Item:=VernalEquinox, Key:=KeyGenerator.Key("Mar")
    
    ' The day for green.
    ' Apr. 29 1989-2006.
    Set GreenApr29 = New SpecialHolidaycls
    GreenApr29.Init ReqStartYear:=1989, ReqEndYear:=2006, ReqMonth:=4, ReqDayofMonth:=29, ReqName:="Ç›Ç«ÇËÇÃì˙"
    mHolidaysCollection.Add Item:=GreenApr29, Key:=KeyGenerator.Key("Apr")
    
    ' May 4 2007-.
    Set GreenMay4 = New SpecialHolidaycls
    GreenMay4.Init ReqStartYear:=2007, ReqMonth:=5, ReqDayofMonth:=4, ReqAltShift:=2, ReqName:="Ç›Ç«ÇËÇÃì˙"
    mHolidaysCollection.Add Item:=GreenMay4, Key:=KeyGenerator.Key("May")
    
    ' The Day of the Showa (Apl. 29 2007-).
    Set DayofShowa = New SpecialHolidaycls
    DayofShowa.Init ReqStartYear:=2007, ReqMonth:=4, ReqDayofMonth:=29, ReqName:="è∫òaÇÃì˙"
    mHolidaysCollection.Add Item:=DayofShowa, Key:=KeyGenerator.Key("Apr")
    
    ' Day of Constitution. Alternate holiday is three days after.
    ' May 3 1948-2006. Alternate holiday is one day after.
    Set Constitution1949 = New SpecialHolidaycls
    Constitution1949.Init ReqStartYear:=1949, ReqEndYear:=2006, ReqMonth:=5, ReqDayofMonth:=3, ReqName:="åõñ@ãLîOì˙"
    mHolidaysCollection.Add Item:=Constitution1949, Key:=KeyGenerator.Key("May")
    
    ' Cildrens' Day
    ' May 5 1949-1987.
    Set Children1949 = New SpecialHolidaycls
    Children1949.Init ReqStartYear:=1949, ReqEndYear:=1987, ReqMonth:=5, ReqDayofMonth:=5, ReqName:="Ç±Ç«Ç‡ÇÃì˙"
    mHolidaysCollection.Add Item:=Children1949, Key:=KeyGenerator.Key("May")
    
    ' May 5 1988-2006, May 4 also goes holiday.
    Set Children1988 = New SpecialHolidaycls
    Children1988.Init ReqStartYear:=1988, ReqEndYear:=2006, ReqMonth:=5, ReqDayofMonth:=5, ReqPrevSHObj:=Constitution1949, ReqName:="Ç±Ç«Ç‡ÇÃì˙"
    mHolidaysCollection.Add Item:=Children1988, Key:=KeyGenerator.Key("May")
    
    ' May 5 2007-.
    Set Children2007 = New SpecialHolidaycls
    Children2007.Init ReqStartYear:=2007, ReqMonth:=5, ReqDayofMonth:=5, ReqName:="Ç±Ç«Ç‡ÇÃì˙"
    mHolidaysCollection.Add Item:=Children2007, Key:=KeyGenerator.Key("May")
    
    ' The Day of Sea
    ' Jul. 20 1996-2002.
    Set DayofSea1996 = New SpecialHolidaycls
    DayofSea1996.Init ReqStartYear:=1996, ReqEndYear:=2002, ReqMonth:=7, ReqDayofMonth:=20, ReqName:="äCÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSea1996, Key:=KeyGenerator.Key("Jul")
    
    ' Third Monday in Jul. 2003-2019.
    Set DayofSea2003 = New SpecialHolidaycls
    DayofSea2003.Init ReqStartYear:=2003, ReqEndYear:=2019, ReqMonth:=7, ReqNumofWeek:=3, ReqName:="äCÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSea2003, Key:=KeyGenerator.Key("Jul")
    
    ' Jul. 23 2020.
    Set DayofSea2020 = New SpecialHolidaycls
    DayofSea2020.Init ReqStartYear:=2020, ReqEndYear:=2020, ReqMonth:=7, ReqDayofMonth:=23, ReqName:="äCÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSea2020, Key:=KeyGenerator.Key("Jul")
    
    ' Jul. 22 2021.
    Set DayofSea2021 = New SpecialHolidaycls
    DayofSea2021.Init ReqStartYear:=2021, ReqEndYear:=2021, ReqMonth:=7, ReqDayofMonth:=22, ReqName:="äCÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSea2021, Key:=KeyGenerator.Key("Jul")
    
    ' Third Monday in Jul. 2022-
    Set DayofSea2022 = New SpecialHolidaycls
    DayofSea2022.Init ReqStartYear:=2022, ReqMonth:=7, ReqNumofWeek:=3, ReqName:="äCÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSea2022, Key:=KeyGenerator.Key("Jul")
    
    ' The day of mountains.
    ' Aug. 11 2016-2019.
    Set DayofMount2016 = New SpecialHolidaycls
    DayofMount2016.Init ReqStartYear:=2016, ReqEndYear:=2019, ReqMonth:=8, ReqDayofMonth:=11, ReqName:="éRÇÃì˙"
    mHolidaysCollection.Add Item:=DayofMount2016, Key:=KeyGenerator.Key("Aug")
    
    ' Aug. 10 2020.
    Set DayofMount2020 = New SpecialHolidaycls
    DayofMount2020.Init ReqStartYear:=2020, ReqEndYear:=2020, ReqMonth:=8, ReqDayofMonth:=10, ReqName:="éRÇÃì˙"
    mHolidaysCollection.Add Item:=DayofMount2020, Key:=KeyGenerator.Key("Aug")

    ' Aug. 8 2021.
    Set DayofMount2021 = New SpecialHolidaycls
    DayofMount2021.Init ReqStartYear:=2021, ReqEndYear:=2021, ReqMonth:=8, ReqDayofMonth:=8, ReqName:="éRÇÃì˙"
    mHolidaysCollection.Add Item:=DayofMount2021, Key:=KeyGenerator.Key("Aug")
    
    ' Aug. 11 2022-.
    Set DayofMount2022 = New SpecialHolidaycls
    DayofMount2022.Init ReqStartYear:=2022, ReqMonth:=8, ReqDayofMonth:=11, ReqName:="éRÇÃì˙"
    mHolidaysCollection.Add Item:=DayofMount2022, Key:=KeyGenerator.Key("Aug")
    
    ' The day of Senior
    ' Sep. 15 1966-2002.
    Set DayofSeniorSep15 = New SpecialHolidaycls
    DayofSeniorSep15.Init ReqStartYear:=1966, ReqEndYear:=2002, ReqMonth:=9, ReqDayofMonth:=15, ReqName:="åhòVÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSeniorSep15, Key:=KeyGenerator.Key("Sep")
    
    ' Third Monday in Sep. 2003-.
    Set DayofSeniorSepTrdMon = New SpecialHolidaycls
    DayofSeniorSepTrdMon.Init ReqStartYear:=2003, ReqMonth:=9, ReqNumofWeek:=3, ReqName:="åhòVÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSeniorSepTrdMon, Key:=KeyGenerator.Key("Sep")
    
    ' The Autumn Equinox
    ' Sep. 1948-1987.
    Set AutumnEquinox1948 = New SpecialHolidaycls
    AutumnEquinox1948.Init ReqEndYear:=1987, ReqIsAutumn_E:=True, ReqName:="èHï™ÇÃì˙"
    mHolidaysCollection.Add Item:=AutumnEquinox1948, Key:=KeyGenerator.Key("Sep")
    
    ' From 1988 additional holiday is set between the day of senior.
    Set AutumnEquinox1988 = New SpecialHolidaycls
    AutumnEquinox1988.Init ReqStartYear:=1988, ReqEndYear:=2002, ReqIsAutumn_E:=True, ReqPrevSHObj:=DayofSeniorSep15, ReqName:="èHï™ÇÃì˙"
    mHolidaysCollection.Add Item:=AutumnEquinox1988, Key:=KeyGenerator.Key("Sep")
    
    Set AutumnEquinox2003 = New SpecialHolidaycls
    AutumnEquinox2003.Init ReqStartYear:=2003, ReqIsAutumn_E:=True, ReqPrevSHObj:=DayofSeniorSepTrdMon, ReqName:="èHï™ÇÃì˙"
    mHolidaysCollection.Add Item:=AutumnEquinox2003, Key:=KeyGenerator.Key("Sep")
    
    ' The Day of Sports.
    ' Oct. 10 1966-1999
    Set DayofSportsOct10 = New SpecialHolidaycls
    DayofSportsOct10.Init ReqStartYear:=1966, ReqEndYear:=1999, ReqMonth:=10, ReqDayofMonth:=10, ReqName:="ëÃàÁÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSportsOct10, Key:=KeyGenerator.Key("Oct")
    
    ' Second Monday in Oct. 2000-2019
    Set DayofSports2000 = New SpecialHolidaycls
    DayofSports2000.Init ReqStartYear:=2000, ReqEndYear:=2019, ReqMonth:=10, ReqNumofWeek:=2, ReqName:="ëÃàÁÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSports2000, Key:=KeyGenerator.Key("Oct")
    
    ' Jul. 24 2020.
    Set DayofSports2020 = New SpecialHolidaycls
    DayofSports2020.Init ReqStartYear:=2020, ReqEndYear:=2020, ReqMonth:=7, ReqDayofMonth:=24, ReqName:="ÉXÉ|Å[ÉcÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSports2020, Key:=KeyGenerator.Key("Jul")
    
    ' Jul. 24 2021.
    Set DayofSports2021 = New SpecialHolidaycls
    DayofSports2021.Init ReqStartYear:=2021, ReqEndYear:=2021, ReqMonth:=7, ReqDayofMonth:=23, ReqName:="ÉXÉ|Å[ÉcÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSports2021, Key:=KeyGenerator.Key("Jul")
    
    ' Second Monday in Oct. 2022-
    Set DayofSportsOct2ndMon = New SpecialHolidaycls
    DayofSportsOct2ndMon.Init ReqStartYear:=2022, ReqMonth:=10, ReqNumofWeek:=2, ReqName:="ÉXÉ|Å[ÉcÇÃì˙"
    mHolidaysCollection.Add Item:=DayofSportsOct2ndMon, Key:=KeyGenerator.Key("Oct")
    
    ' Day of Culture (Nov. 3)
    Set Culture = New SpecialHolidaycls
    Culture.Init ReqMonth:=11, ReqDayofMonth:=3, ReqName:="ï∂âªÇÃì˙"
    mHolidaysCollection.Add Item:=Culture, Key:=KeyGenerator.Key("Nov")
    
    ' Appliciation day for work (Nov. 23)
    Set Work = New SpecialHolidaycls
    Work.Init ReqMonth:=11, ReqDayofMonth:=23, ReqName:="ãŒòJä¥é”ÇÃì˙"
    mHolidaysCollection.Add Item:=Work, Key:=KeyGenerator.Key("Nov")
    
    ' Birthday of the Emperor.
    ' Apr. 29.
    Set EmperorApr29 = New SpecialHolidaycls
    EmperorApr29.Init ReqStartYear:=1949, ReqEndYear:=1988, ReqMonth:=4, ReqDayofMonth:=29, ReqName:="ìVçcíaê∂ì˙"
    mHolidaysCollection.Add Item:=EmperorApr29, Key:=KeyGenerator.Key("Apr")
    
    ' Dec. 23.
    Set EmperorDec23 = New SpecialHolidaycls
    EmperorDec23.Init ReqStartYear:=1989, ReqEndYear:=2018, ReqMonth:=12, ReqDayofMonth:=23, ReqName:="ìVçcíaê∂ì˙"
    mHolidaysCollection.Add Item:=EmperorDec23, Key:=KeyGenerator.Key("Dec")
    
    ' Feb. 23.
    Set EmperorFeb23 = New SpecialHolidaycls
    EmperorFeb23.Init ReqStartYear:=2020, ReqMonth:=2, ReqDayofMonth:=23, ReqName:="ìVçcíaê∂ì˙"
    mHolidaysCollection.Add Item:=EmperorFeb23, Key:=KeyGenerator.Key("Feb")
    
    ' The marriage of the prince in 1959.
    Set MarriageofPrince1959 = New SpecialHolidaycls
    MarriageofPrince1959.Init ReqStartYear:=1959, ReqEndYear:=1959, ReqMonth:=4, ReqDayofMonth:=10, ReqName:="çcëæéqåãç•ÇÃãV"
    mHolidaysCollection.Add Item:=MarriageofPrince1959, Key:=KeyGenerator.Key("Apr")
    
    ' The funeral of the emperor in 1989.
    Set FuneralofEmperor1989 = New SpecialHolidaycls
    FuneralofEmperor1989.Init ReqStartYear:=1989, ReqEndYear:=1989, ReqMonth:=2, ReqDayofMonth:=24, ReqName:="ëÂërÇÃóÁ"
    mHolidaysCollection.Add Item:=FuneralofEmperor1989, Key:=KeyGenerator.Key("Feb")
    
    ' The enthronemento of the emperor in 1990.
    Set EnthronementofEmperor1990 = New SpecialHolidaycls
    EnthronementofEmperor1990.Init ReqStartYear:=1990, ReqEndYear:=1990, ReqMonth:=11, ReqDayofMonth:=12, ReqName:="ë¶à óÁê≥ìaÇÃãV"
    mHolidaysCollection.Add Item:=EnthronementofEmperor1990, Key:=KeyGenerator.Key("Nov")
    
    ' The marriage of prince in 1993.
    Set MarriageofPrince1993 = New SpecialHolidaycls
    MarriageofPrince1993.Init ReqStartYear:=1993, ReqEndYear:=1993, ReqMonth:=6, ReqDayofMonth:=9, ReqName:="çcëæéqåãç•ÇÃãV"
    mHolidaysCollection.Add Item:=MarriageofPrince1993, Key:=KeyGenerator.Key("Jun")
    
    ' The enthronemento of emperor in 2019.
    Set EnthronementofEmperor2019 = New SpecialHolidaycls
    EnthronementofEmperor2019.Init ReqStartYear:=2019, ReqEndYear:=2019, ReqMonth:=5, ReqDayofMonth:=1, ReqPrevSHObj:=DayofShowa, ReqName:="ìVçcÇÃë¶à ÇÃì˙"
    mHolidaysCollection.Add Item:=EnthronementofEmperor2019, Key:=KeyGenerator.Key("May")
        
    ' The enthronemento ceremony in 2019.
    Set EnthronementCeremony2019 = New SpecialHolidaycls
    EnthronementCeremony2019.Init ReqStartYear:=2019, ReqEndYear:=2019, ReqMonth:=10, ReqDayofMonth:=22, ReqName:="ë¶à óÁê≥ìaÇÃãV"
    mHolidaysCollection.Add Item:=EnthronementCeremony2019, Key:=KeyGenerator.Key("Oct")
    
    ' The day of constitution.
    ' May 3 2007-. Alternate holiday is three day after.
    Set Constitution2007 = New SpecialHolidaycls
    Constitution2007.Init ReqStartYear:=2007, ReqMonth:=5, ReqDayofMonth:=3, ReqAltShift:=3, ReqPrevSHObj:=EnthronementofEmperor2019, ReqName:="åõñ@ãLîOì˙"
    mHolidaysCollection.Add Item:=Constitution2007, Key:=KeyGenerator.Key("May")

End Sub
