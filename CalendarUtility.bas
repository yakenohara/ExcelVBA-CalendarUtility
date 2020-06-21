Attribute VB_Name = "CalendarUtility"
'<License>------------------------------------------------------------
'
' Copyright (c) 2020 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
' カレンダー上の指定日から起算して、指定時間を消費し終わる日を算出して返す
'
' Parameters
' ----------
' startDay : Date
'   起算開始日
'   時刻が含まれる場合は、その時刻から起算開始する
'
' hours : Double
'   消費する時間数の合計[h]
'
' calendar : Range or String
'   日毎の消費時間数定義表。以下の条件を満たす定義であること
'    - 1列目に Date 型の日付、2列目にその日に消費する時間数[h](Double型にキャスト可能な数値)が定義されていること
'    - 1行毎に1日が定義されており、日付の省略行が存在しないこと(e.g. 1月/1日の次の行が1月/3日であってはならない
'
' targetDay : Date, Default Nothing
'   指定された場合、その日までに消費する時間数[h] を返す
'
' integral : Boolean, Default True
'   targetDayが指定された場合にのみ有効
'   True が 指定された場合、targetDay の日までに消費する時間数[h] を返す
'   False が 指定された場合、targetDay の日に消費する時間数[h] を返す
'
Public Function calendarHowSpend( _
    ByVal startDay As Date, _
    ByVal hours As Double, _
    ByVal calendar As Variant, _
    Optional ByVal targetDay As Variant = Nothing, _
    Optional ByVal integral As Boolean = True _
) As Variant

    Dim date_startDay As Date
    Dim date_endDay As Date
    Dim date_targetDay As Variant
    Dim dbl_startHour As Double
    Dim long_rowOfStartDay As Long
    Dim vararr_calendar As Variant
    
    Dim int_dateCol As Integer
    Dim int_hourCol As Integer
    
    Dim expendedHoursPerDay As Double
    Dim expendedHours As Double
    Dim long_spendDays As Long
    Dim dbl_remainInDay As Double
    
    Dim toRet As Variant
    
    date_startDay = (Year(startDay) & "/" & Month(startDay) & "/" & Day(startDay)) '時間を破棄して日付を取得
    dbl_startHour = (hour(startDay) + Minute(startDay) / 60)                       '日付を破棄して時間を取得

    Set date_targetDay = Nothing
    If TypeName(targetDay) = "Range" Then ' Range 型の場合
    
        If TypeName(targetDay.Value) <> "Date" Then ' Range が表すセル内の値の型が Date 型ではない場合
            toRet = xlErrValue ' #VALUE! を返す
            GoTo RETURN_AND_EXIT
            
        End If
        
        date_targetDay = targetDay.Value
        
    ElseIf TypeName(targetDay) = "Date" Then ' Date 型の場合
        date_targetDay = targetDay
        
    End If
    
    str_tmp = TypeName(date_targetDay)
    If (str_tmp <> "Date" And str_tmp <> "Nothing") Then
        toRet = xlErrValue ' #VALUE! を返す
        GoTo RETURN_AND_EXIT

    End If
    
    If TypeName(date_targetDay) = "Date" Then
        If (date_targetDay < date_startDay) Then
            toRet = 0 ' 0 を返す
            GoTo RETURN_AND_EXIT
        End If
    End If
    
    bool_foundStartDay = False
    
    If TypeName(calendar) = "String" Then ' 日毎の消費時間数定義表が String 型で指定された場合
        vararr_calendar = Range(calendar)
        
    Else ' 日毎の消費時間数定義表が String 型以外の場合
        vararr_calendar = calendar ' Range オブジェクトとみなす
        
    End If
    
    
    '開始日をカレンダーから検索する
    For long_rowOfStartDay = LBound(vararr_calendar, 1) To UBound(vararr_calendar, 1)
        
        int_dateCol = LBound(vararr_calendar, 2) '一番左の列を日付列とみなす
        int_hourCol = int_dateCol + 1            '一番左 + 1 の列を時間定義列とみなす
        
        If (vararr_calendar(long_rowOfStartDay, int_dateCol) = date_startDay) Then '開始日が見つかった時
            bool_foundStartDay = True
            Exit For
        End If
        
    Next long_rowOfStartDay
    
    If (Not bool_foundStartDay) Then '開始日が見つからなかった時
        toRet = xlErrValue ' #VALUE! を返す
        
    Else '開始日が見つかった時
        
        expendedHoursPerDay = 0 'その日に消費する時間数
        expendedHours = 0 '積算消費時間
        long_spendDays = 0  '消費日数(0 based)
        
        
        'タスク終了日か指定日まで日毎時間数を積算するループ
        bl_isFirst = True
        Do While True
            
            If bl_isFirst Then ' ループの一回目の場合
                
                '日毎の残り時間(消費可能時間)の算出
                dbl_remainInDay = (vararr_calendar(long_rowOfStartDay + long_spendDays, int_hourCol) * 1) - dbl_startHour
                
                ' 最初の日の開始時間がその日の時間数を超えている場合
                ' e.g. 最初の日の開始時間が 9:00 だけど、その日の時間数は 8.00 [h]
                ' ただし、0 [h] は許容する。次の日から積算開始する事と同義だから。
                ' e.g. 最初の日の開始時間が 8:00 だけど、その日の時間数は 8.00 [h]
                '      -> これは次の日の 0:00 から積算開始する事と同じ。
                If dbl_remainInDay < 0 Then
                    toRet = xlErrValue ' #VALUE! を返す
                    GoTo RETURN_AND_EXIT
                End If
                
                bl_isFirst = False
                
            Else
                'カレンダーが 1日 一行となっているかどうか確認
                int_expectedAs1Day = DateDiff("d", vararr_calendar(long_rowOfStartDay + long_spendDays - 1, int_dateCol), vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol))
                If (int_expectedAs1Day <> 1) Then 'カレンダーが 1日 一行となっていない場合
                    toRet = xlErrValue ' #VALUE! を返す
                    GoTo RETURN_AND_EXIT
                End If
                
                '日毎の残り時間(消費可能時間)の算出
                dbl_remainInDay = (vararr_calendar(long_rowOfStartDay + long_spendDays, int_hourCol) * 1)
            
            End If
            
            bool_tmp = False
            If TypeName(date_targetDay) = "Date" Then
                bool_tmp = (vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol) = date_targetDay)
            End If
            
            
            'その日でタスクの時間をすべて消費し終わる場合
            If (hours <= (expendedHours + dbl_remainInDay)) Then
                
                long_flacHour = 0
                If long_spendDays = 0 Then ' ループの一回目の場合
                    long_flacHour = dbl_startHour '算出する終了時間にあらかじめ開始時間を足しておく
                End If
                
                long_flacHour = long_flacHour + (hours - expendedHours) '終了時間の算出
                expendedHoursPerDay = hours - expendedHours
                expendedHours = hours '積算消費時間はタスクの時間数に等しい
                
                Exit Do
                
            ElseIf bool_tmp Then '指定日の場合
                
                long_flacHour = 0
                If long_spendDays = 0 Then ' ループの一回目の場合
                    long_flacHour = dbl_startHour '算出する終了時間にあらかじめ開始時間を足しておく
                End If
                
                long_flacHour = long_flacHour + (hours - expendedHours) '終了時間の算出
                expendedHoursPerDay = dbl_remainInDay
                expendedHours = expendedHours + dbl_remainInDay '積算消費時間はタスクの時間数に等しい
                
                Exit Do
                
            'その日でタスクの時間をすべて消費し終わらない And
            '指定日でもない場合
            Else
                expendedHoursPerDay = dbl_remainInDay
                expendedHours = expendedHours + dbl_remainInDay '積算消費時間にその日の消費可能時間を和算
                long_spendDays = long_spendDays + 1 '消費日数(0 based) += 1
                
            End If
            
        Loop
        
        If TypeName(date_targetDay) = "Date" Then
            
            If integral Then '積算消費時間指定の場合
                toRet = expendedHours
                
            Else  '指定日の消費時間指定の場合
            
                date_endDay = vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol)
                
                If date_endDay < date_targetDay Then
                    toRet = 0
                    
                Else
                    toRet = expendedHoursPerDay
                    
                End If
                
                'todo 開始日を下回る
                
            End If

        Else
            date_endDay = vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol)
            date_endDay = DateAdd("h", long_flacHour, date_endDay)
            date_endDay = DateAdd("n", ((long_flacHour - Fix(long_flacHour)) * 60), date_endDay)
            toRet = date_endDay

        End If
        
        
    End If
    
RETURN_AND_EXIT:
    
    calendarHowSpend = toRet
    
End Function

'
' カレンダー上の指定日時から起算して、微小時間後の日時を返す
' e.g. 1月1日の消費可能時間[h] が 8[h] とカレンダーに定義されていて、
'      指定日時が "1/1 8:00" だったときは、"1/2 0:00" を返す
'
' Parameters
' ----------
' dateToAdd : Date
'   起算開始日
'   時刻が含まれる場合は、その時刻から起算開始する
'
' calendar : Range or String
'   日毎の消費時間数定義表。以下の条件を満たす定義であること
'    - 1列目に Date 型の日付、2列目にその日に消費する時間数[h](Double型にキャスト可能な数値)が定義されていること
'    - 1行毎に1日が定義されており、日付の省略行が存在しないこと(e.g. 1月/1日の次の行が1月/3日であってはならない
'
Public Function calendarAddDelta( _
    ByVal dateToAdd As Date, _
    ByVal calendar As Variant _
) As Variant
    
    Dim date_startDay As Date
    Dim dbl_startHour As Double
    Dim long_rowOfStartDay As Long
    Dim vararr_calendar As Variant
    
    Dim int_dateCol As Integer
    Dim int_hourCol As Integer

    Dim long_spendDays As Long
    Dim dbl_remainInDay As Double
    
    date_startDay = (Year(dateToAdd) & "/" & Month(dateToAdd) & "/" & Day(dateToAdd)) '時間を破棄して日付を取得
    dbl_startHour = (hour(dateToAdd) + Minute(dateToAdd) / 60)                        '日付を破棄して時間を取得
    
    bool_foundStartDay = False
    
    If TypeName(calendar) = "String" Then ' 日毎の消費時間数定義表が String 型で指定された場合
        vararr_calendar = Range(calendar)
        
    Else ' 日毎の消費時間数定義表が String 型以外の場合
        vararr_calendar = calendar ' Range オブジェクトとみなす
        
    End If
    
    '起算開始日をカレンダーから検索する
    For long_rowOfStartDay = LBound(vararr_calendar, 1) To UBound(vararr_calendar, 1)
        
        int_dateCol = LBound(vararr_calendar, 2) '一番左の列を日付列とみなす
        int_hourCol = int_dateCol + 1            '一番左 + 1 の列を時間定義列とみなす
        
        If (vararr_calendar(long_rowOfStartDay, int_dateCol) = date_startDay) Then '起算開始日が見つかった時
            bool_foundStartDay = True
            Exit For
        End If
        
    Next long_rowOfStartDay
    
    If (Not bool_foundStartDay) Then '起算開始日が見つからなかった時
        toRet = xlErrValue ' #VALUE! を返す
        
    Else '起算開始日が見つかった時
        
        dbl_remainInDay = vararr_calendar(long_rowOfStartDay, int_hourCol) - dbl_startHour

        ' 起算開始日の時間がその日の時間数を超えている場合
        ' e.g. 起算開始日の開始時間が 9:00 だけど、その日の時間数は 8.00 [h]
        If dbl_remainInDay < 0 Then
            Err.Raise (xlErrValue)  ' 強制的に #VALUE! を返す

        ElseIf dbl_remainInDay = 0 Then ' 起算開始日の時間がその日の時間数と一致する場合

            '次の開始日を検索する
            long_spendDays = 0
            Do While True

                long_spendDays = long_spendDays + 1
                
                If 0 < (vararr_calendar(long_rowOfStartDay + long_spendDays, int_hourCol) * 1) Then
                    toRet = vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol)
                    Exit Do
                End If

            Loop

        Else  ' 起算開始日の時間がその日の時間数より少ない場合
            toRet = dateToAdd

        End If
        
    End If
    
    calendarAddDelta = toRet
    
End Function

