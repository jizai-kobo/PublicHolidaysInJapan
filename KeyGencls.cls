VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyGencls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private mJanKeyCol As Collection
Private mFebKeyCol As Collection
Private mMarKeyCol As Collection
Private mAprKeyCol As Collection
Private mMayKeyCol As Collection
Private mJunKeyCol As Collection
Private mJulKeyCol As Collection
Private mAugKeyCol As Collection
Private mSepKeyCol As Collection
Private mOctKeyCol As Collection
Private mNovKeyCol As Collection
Private mDecKeyCol As Collection

' Key phrase generating function.
Public Function Key(ByVal ReqMonthStr As String) As String
    Select Case ReqMonthStr
        Case "Jan"
            mJanKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mJanKeyCol.Count, "00")
            Exit Function
        Case "Feb"
            mFebKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mFebKeyCol.Count, "00")
            Exit Function
        Case "Mar"
            mMarKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mMarKeyCol.Count, "00")
            Exit Function
        Case "Apr"
            mAprKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mAprKeyCol.Count, "00")
            Exit Function
        Case "May"
            mMayKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mMayKeyCol.Count, "00")
            Exit Function
        Case "Jun"
            mJunKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mJunKeyCol.Count, "00")
            Exit Function
        Case "Jul"
            mJulKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mJulKeyCol.Count, "00")
            Exit Function
        Case "Aug"
            mAugKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mAugKeyCol.Count, "00")
            Exit Function
        Case "Sep"
            mSepKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mSepKeyCol.Count, "00")
            Exit Function
        Case "Oct"
            mOctKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mOctKeyCol.Count, "00")
            Exit Function
        Case "Nov"
            mNovKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mNovKeyCol.Count, "00")
            Exit Function
        Case "Dec"
            mDecKeyCol.Add ReqMonthStr
            Key = ReqMonthStr & Format(mDecKeyCol.Count, "00")
            Exit Function
        Case Else
            Err.Raise 901, Source:="祝日発生ルーチン", Description:="エラー：キーの設定でエラーが発生しました。"
    End Select
End Function

' Constructor.
Private Sub Class_Initialize()
    Set mJanKeyCol = New Collection
    Set mFebKeyCol = New Collection
    Set mMarKeyCol = New Collection
    Set mAprKeyCol = New Collection
    Set mMayKeyCol = New Collection
    Set mJunKeyCol = New Collection
    Set mJulKeyCol = New Collection
    Set mAugKeyCol = New Collection
    Set mSepKeyCol = New Collection
    Set mOctKeyCol = New Collection
    Set mNovKeyCol = New Collection
    Set mDecKeyCol = New Collection
End Sub
