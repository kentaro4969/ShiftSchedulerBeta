' ShiftSchedulerBeta.bas
' Version: Beta1.0
' Author: Kentaro Nakamura
' License: Non-Commercial Use Only
' 
' This VBA script is for generating and managing shift schedules.
' This script is provided under the following terms:
'
' 1. Non-Commercial Use Only: This script may be used for personal, educational, or research purposes only. Commercial use is strictly prohibited.
' 2. Modification: You are free to modify the script for your own use, but you may not distribute modified versions for commercial purposes.
' 3. Distribution: You may distribute copies of the original script as long as this license file is included and the script is not used for commercial purposes.
' 4. No Warranty: This script is provided "as is", without warranty of any kind.
'
' By using this script, you agree to abide by these terms.

' ShiftMarksというユーザー定義型を宣言
Type ShiftMarks
    dayShiftMark As String
    nightShiftMark As String
    nightShiftAfterMark As String
    holidayMark As String
End Type

' グローバル変数を宣言
Dim g_marks As ShiftMarks
Dim g_staffCompatibility() As String
Dim g_consecutiveNightShiftAbility() As String
Dim staffFridaySaturdayOnlyNightShift() As String 
Dim staffMaxHours() As Double
Dim staffDayShiftMax() As Integer
Dim staffNightShiftMax() As Integer
Dim previousMonthLastNightShift() As Integer ' 前月末の夜勤情報を保持
Dim previousMonthConsecutiveDayShift() As Integer ' 前月末の連続日勤数を追加
Dim totalShiftsAssigned() As Integer ' 総シフト数を追跡
Dim g_totalDates As Integer
Dim staffHolidays() As Collection
Dim staffMaxConsecutiveDayShifts() As Integer 

Sub GenerateShift()
    On Error GoTo ErrorHandler

    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim hospitalName As String
    Dim startDate As Date
    Dim weekdayDayShiftMax As Integer
    Dim weekdayDayShiftMin As Integer
    Dim weekdayNightShiftMax As Integer
    Dim weekdayNightShiftMin As Integer
    Dim weekendDayShiftMax As Integer
    Dim weekendDayShiftMin As Integer
    Dim weekendNightShiftMax As Integer
    Dim weekendNightShiftMin As Integer
    Dim dayShiftLeaderMin As Integer
    Dim nightShiftLeaderMin As Integer
    Dim staffNames() As String
    Dim staffAttributes() As String
    Dim staffShifts() As Collection
    Dim i As Integer, j As Integer, k As Integer
    Dim currentDate As Date
    Dim dayOfWeek As Integer
    Dim dayShiftCount As Integer
    Dim nightShiftCount As Integer
    Dim hoursWorked() As Double
    Dim nightShiftCountWorked() As Integer
    Dim availableStaff As Collection
    Dim staffIndex As Variant
    Dim logFile As String
    Dim holidayDates As Collection
    Dim holidayDate As Variant

    logFile = ThisWorkbook.Path & "\ShiftScheduleLog.txt"
    Open logFile For Output As #1
    Print #1, "Step 1: Setting worksheet"

    Set wsInput = ThisWorkbook.Sheets("Input") ' 入力シートからデータを取得

    Print #1, "Step 2: Getting input data"
    On Error GoTo DataError

    hospitalName = wsInput.Cells(1, 2).Value
    Print #1, "hospitalName: " & hospitalName
    startDate = wsInput.Cells(2, 2).Value
    Print #1, "startDate: " & startDate
    weekdayDayShiftMax = wsInput.Cells(3, 2).Value
    Print #1, "weekdayDayShiftMax: " & weekdayDayShiftMax
    weekdayDayShiftMin = wsInput.Cells(3, 3).Value
    Print #1, "weekdayDayShiftMin: " & weekdayDayShiftMin
    weekdayNightShiftMax = wsInput.Cells(4, 2).Value
    Print #1, "weekdayNightShiftMax: " & weekdayNightShiftMax
    weekdayNightShiftMin = wsInput.Cells(4, 3).Value
    Print #1, "weekdayNightShiftMin: " & weekdayNightShiftMin
    weekendDayShiftMax = wsInput.Cells(5, 2).Value
    Print #1, "weekendDayShiftMax: " & weekendDayShiftMax
    weekendDayShiftMin = wsInput.Cells(5, 3).Value
    Print #1, "weekendDayShiftMin: " & weekendDayShiftMin
    weekendNightShiftMax = wsInput.Cells(6, 2).Value
    Print #1, "weekendNightShiftMax: " & weekendNightShiftMax
    weekendNightShiftMin = wsInput.Cells(6, 3).Value
    Print #1, "weekendNightShiftMin: " & weekendNightShiftMin
    dayShiftLeaderMin = wsInput.Cells(7, 2).Value
    Print #1, "dayShiftLeaderMin: " & dayShiftLeaderMin
    nightShiftLeaderMin = wsInput.Cells(8, 2).Value
    Print #1, "nightShiftLeaderMin: " & nightShiftLeaderMin

    ' シフトマークを取得
    g_marks.dayShiftMark = wsInput.Cells(1, 16).Value ' P1
    g_marks.nightShiftMark = wsInput.Cells(2, 16).Value ' P2
    g_marks.nightShiftAfterMark = wsInput.Cells(3, 16).Value ' P3
    g_marks.holidayMark = wsInput.Cells(4, 16).Value ' P4

    On Error GoTo ErrorHandler

    ' スタッフ情報を取得
    Print #1, "Step 3: Getting staff data"
    On Error GoTo StaffDataError

    Dim lastRow As Integer
    lastRow = wsInput.Cells(wsInput.Rows.count, 1).End(xlUp).Row
    Print #1, "lastRow: " & lastRow

    ReDim staffNames(1 To lastRow - 9)
    ReDim staffAttributes(1 To lastRow - 9)
    ReDim staffMaxHours(1 To lastRow - 9)
    ReDim staffDayShiftMax(1 To lastRow - 9)
    ReDim staffNightShiftMax(1 To lastRow - 9)
    ReDim g_staffCompatibility(1 To lastRow - 9)
    ReDim staffShifts(1 To lastRow - 9)
    ReDim staffHolidays(1 To lastRow - 9)
    ReDim hoursWorked(1 To lastRow - 9)
    ReDim nightShiftCountWorked(1 To lastRow - 9)
    ReDim previousMonthLastNightShift(1 To lastRow - 9)
    ReDim previousMonthConsecutiveDayShift(1 To lastRow - 9)
    ReDim totalShiftsAssigned(1 To lastRow - 9)
    ReDim staffFridaySaturdayOnlyNightShift(1 To lastRow - 9) ' 追加
    ReDim staffMaxConsecutiveDayShifts(1 To lastRow - 9)

    Dim staffCount As Integer
    staffCount = 0

    For i = 10 To lastRow
        If IsEmpty(wsInput.Cells(i, 1).Value) Then
            Print #1, "Skipping empty staff name at row: " & i
            GoTo NextStaff
        End If

        staffCount = staffCount + 1

        ReDim Preserve staffNames(1 To staffCount)
        ReDim Preserve staffAttributes(1 To staffCount)
        ReDim Preserve staffMaxHours(1 To staffCount)
        ReDim Preserve staffDayShiftMax(1 To staffCount)
        ReDim Preserve staffNightShiftMax(1 To staffCount)
        ReDim Preserve g_staffCompatibility(1 To staffCount)
        ReDim Preserve staffShifts(1 To staffCount)
        ReDim Preserve staffHolidays(1 To staffCount)
        ReDim Preserve hoursWorked(1 To staffCount)
        ReDim Preserve nightShiftCountWorked(1 To staffCount)
        ReDim Preserve previousMonthLastNightShift(1 To staffCount)
        ReDim Preserve previousMonthConsecutiveDayShift(1 To staffCount)
        ReDim Preserve totalShiftsAssigned(1 To staffCount)
        ReDim Preserve staffFridaySaturdayOnlyNightShift(1 To staffCount) ' 追加
        ReDim Preserve staffMaxConsecutiveDayShifts(1 To staffCount)

' F列から最大連続日勤日数を取得
staffMaxConsecutiveDayShifts(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 6).Value), wsInput.Cells(i, 6).Value, 5) ' F列
Print #1, "Staff " & staffNames(staffCount) & " max consecutive day shifts: " & staffMaxConsecutiveDayShifts(staffCount)


        staffNames(staffCount) = wsInput.Cells(i, 1).Value
        If IsEmpty(wsInput.Cells(i, 2).Value) Then
            staffAttributes(staffCount) = ""
        Else
            staffAttributes(staffCount) = wsInput.Cells(i, 2).Value
        End If

        ' 相性データを取得
        g_staffCompatibility(staffCount) = Trim(wsInput.Cells(i, 4).Value)
        Print #1, "Staff " & staffNames(staffCount) & " compatibility: " & g_staffCompatibility(staffCount)

        ' 連続夜勤適性を3列目から取得
        ReDim Preserve g_consecutiveNightShiftAbility(1 To staffCount)
        g_consecutiveNightShiftAbility(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 3).Value), wsInput.Cells(i, 3).Value, "")
        Print #1, "Staff " & staffNames(staffCount) & " consecutive night shift ability: " & g_consecutiveNightShiftAbility(staffCount)

        ' 前月末の夜勤情報を取得（21列目）
        previousMonthLastNightShift(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 21).Value), wsInput.Cells(i, 21).Value, 0)
        Print #1, "Staff " & staffNames(staffCount) & " previous month last night shift: " & previousMonthLastNightShift(staffCount)

        ' 前月末の連続日勤数を取得（22列目）
        previousMonthConsecutiveDayShift(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 22).Value), wsInput.Cells(i, 22).Value, 0)
        Print #1, "Staff " & staffNames(staffCount) & " previous month consecutive day shifts: " & previousMonthConsecutiveDayShift(staffCount)

        staffMaxHours(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 8).Value), wsInput.Cells(i, 8).Value, 0) ' H列
        staffDayShiftMax(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 9).Value), wsInput.Cells(i, 9).Value, 0) ' I列
        staffNightShiftMax(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 10).Value), wsInput.Cells(i, 10).Value, 0) ' J列

        ' 属性を取得（E列）
        If IsEmpty(wsInput.Cells(i, 5).Value) Then
            staffFridaySaturdayOnlyNightShift(staffCount) = ""
        Else
            staffFridaySaturdayOnlyNightShift(staffCount) = wsInput.Cells(i, 5).Value
        End If
        Print #1, "Staff " & staffNames(staffCount) & " Friday/Saturday only night shift: " & staffFridaySaturdayOnlyNightShift(staffCount)

        Set staffShifts(staffCount) = New Collection
        Set staffHolidays(staffCount) = New Collection

        Print #1, "Staff: " & staffNames(staffCount) & ", Attribute: " & staffAttributes(staffCount) & _
        ", Compatibility: " & g_staffCompatibility(staffCount) & ", MaxHours: " & staffMaxHours(staffCount) & _
        ", DayShiftMax: " & staffDayShiftMax(staffCount) & ", NightShiftMax: " & staffNightShiftMax(staffCount)

        For j = 11 To 20 
            If IsDate(wsInput.Cells(i, j).Value) Then
                staffHolidays(staffCount).Add wsInput.Cells(i, j).Value
                Print #1, " Holiday: " & wsInput.Cells(i, j).Value
            Else
                Print #1, " Non-date value ignored: " & wsInput.Cells(i, j).Value
            End If
        Next j

NextStaff:
    Next i

    On Error GoTo ErrorHandler

' 公休日を取得
Print #1, "Step 4: Getting public holidays"
On Error GoTo HolidayDataError

Set holidayDates = New Collection
' 公休日が入力される列を明示的に指定（列E（5）から列J（10）まで）
For colIndex = 5 To 10 ' 列Eから列J
    If IsDate(wsInput.Cells(2, colIndex).Value) Then
        holidayDates.Add wsInput.Cells(2, colIndex).Value
        Print #1, " Public Holiday: " & wsInput.Cells(2, colIndex).Value
    End If
Next colIndex

    On Error GoTo ErrorHandler

    ' 出力シートの作成
    Print #1, "Step 5: Creating output sheet"
    On Error GoTo OutputSheetError

    Dim sheetName As String
    sheetName = "Shift_" & Format(Now, "yyyyMMdd_HHmmss")
    Set wsOutput = ThisWorkbook.Sheets.Add
    wsOutput.Name = sheetName

    On Error GoTo ErrorHandler

    ' ヘッダー行を作成
    Print #1, "Step 6: Creating header row"
    wsOutput.Cells(1, 1).Value = hospitalName
    wsOutput.Cells(1, 2).Value = Format(startDate, "yyyy-MM-dd")
    wsOutput.Cells(1, 3).Value = "Generated on: " & Format(Now, "yyyy-MM-dd HH:mm:ss")

    ' 日付の行を作成
    Print #1, "Step 7: Creating date row"
    On Error GoTo DateRowError

    Dim endDate As Date
    endDate = DateAdd("m", 1, startDate)
    endDate = DateAdd("d", -1, endDate) ' シフト表が1ヶ月分のみ作成されるように修正
    Dim shiftDates As Collection
    Set shiftDates = New Collection

    currentDate = startDate
    Do While currentDate <= endDate
        shiftDates.Add currentDate
        wsOutput.Cells(2, shiftDates.count + 2).Value = Format(currentDate, "m/d(aaa)")
        currentDate = DateAdd("d", 1, currentDate)
    Loop

    ' 日付の総数をグローバル変数に保存
    g_totalDates = shiftDates.count

    On Error GoTo ErrorHandler

    ' スタッフ名と属性の列を作成
    Print #1, "Step 8: Creating staff name and attribute columns"
    On Error GoTo StaffNameColumnError

    For i = 1 To UBound(staffNames)
        If staffNames(i) <> "" Then
            wsOutput.Cells(i + 2, 1).Value = staffNames(i)
            wsOutput.Cells(i + 2, 2).Value = staffAttributes(i)
        End If
    Next i

    On Error GoTo ErrorHandler

    ' セルの書式を「縮小して全体を表示する」に設定
    wsOutput.Cells.Style.ShrinkToFit = True

    ' 指定休日を休みに設定し、セルの色を変更
    Print #1, "Step 9: Setting holidays"
    On Error GoTo HolidaySettingError

    For i = 1 To UBound(staffNames)
        If staffNames(i) <> "" Then
            For Each holidayDate In staffHolidays(i)
                For k = 1 To shiftDates.count
                    If shiftDates(k) = holidayDate Then
                        wsOutput.Cells(i + 2, k + 2).Value = g_marks.holidayMark
                        wsOutput.Cells(i + 2, k + 2).Interior.Color = RGB(255, 255, 153) ' 薄い黄色
                        Exit For
                    End If
                Next k
            Next holidayDate
        End If
    Next i

    ' 公休日のセルの色を変更
    For k = 1 To shiftDates.count
        currentDate = shiftDates(k)
        dayOfWeek = Weekday(currentDate, vbSunday)
        If dayOfWeek = vbSaturday Then
            wsOutput.Cells(2, k + 2).Interior.Color = RGB(173, 216, 230) ' 薄いブルー
        ElseIf dayOfWeek = vbSunday Then
            wsOutput.Cells(2, k + 2).Interior.Color = RGB(255, 182, 193) ' 薄いピンク
        End If

        For Each holidayDate In holidayDates
            If currentDate = holidayDate Then
                wsOutput.Cells(2, k + 2).Interior.Color = RGB(255, 182, 193) ' 薄いピンク
            End If
        Next holidayDate
    Next k

    On Error GoTo ErrorHandler

    ' 追加: 前月末の夜勤データを処理するステップ
    Print #1, "Step 9b: Processing previous month's last night shift data"
    For i = 1 To UBound(staffNames)
        If previousMonthLastNightShift(i) = 1 Then ' 前月末が夜勤の場合
            ' 当月1日は夜勤明け、2日は休み
            If wsOutput.Cells(i + 2, 3).Value = "" Then
                wsOutput.Cells(i + 2, 3).Value = g_marks.nightShiftAfterMark ' 1日目
            End If
            If wsOutput.Cells(i + 2, 4).Value = "" Then
                wsOutput.Cells(i + 2, 4).Value = g_marks.holidayMark ' 2日目
            End If
        ElseIf previousMonthLastNightShift(i) = 2 Then ' 前月末が夜勤明けの場合
            ' 当月1日は休み
            If wsOutput.Cells(i + 2, 3).Value = "" Then
                wsOutput.Cells(i + 2, 3).Value = g_marks.holidayMark ' 1日目
            End If
        End If
    Next i

    ' 追加: 前月末の連続日勤数を考慮する
    Print #1, "Step 9c: Processing previous month's consecutive day shifts"
    For i = 1 To UBound(staffNames)
        If previousMonthConsecutiveDayShift(i) >= 5 Then
            ' 前月末の連続日勤が5日以上の場合、当月1日は休み
            If wsOutput.Cells(i + 2, 3).Value = "" Then
                wsOutput.Cells(i + 2, 3).Value = g_marks.holidayMark
                Print #1, "Staff " & staffNames(i) & " had " & previousMonthConsecutiveDayShift(i) & " consecutive day shifts last month, assigning day off on first day"
            End If
        End If
    Next i

    ' 初期化
    ReDim hoursWorked(1 To staffCount)
    ReDim nightShiftCountWorked(1 To staffCount)
    ReDim totalShiftsAssigned(1 To staffCount)

    ' 日毎のシフトカウント用変数を初期化
    Dim dayShiftDailyCount() As Integer
    Dim nightShiftDailyCount() As Integer
    Dim offShiftDailyCount() As Integer
    ReDim dayShiftDailyCount(1 To shiftDates.count)
    ReDim nightShiftDailyCount(1 To shiftDates.count)
    ReDim offShiftDailyCount(1 To shiftDates.count)

    ' 各ステップを順番に全日付に対して実行

    ' ステップ1: 管理者を全ての日勤に割り付ける
    Print #1, "Step 10: Assigning administrators to day shifts"
    For i = 1 To shiftDates.count
        currentDate = shiftDates(i)
        dayOfWeek = Weekday(currentDate, vbSunday)
        ' 管理者の属性が"4"の場合は休日（日曜または祝日）には日勤を割り当てない
        For j = 1 To UBound(staffNames)
            If staffAttributes(j) = "4" Then
                ' 土曜、日曜、または祝日の場合は日勤をスキップ
                If dayOfWeek = vbSaturday Or dayOfWeek = vbSunday Or IsHoliday(currentDate, holidayDates) Then
                    Print #1, "Skipping day shift assignment for administrator " & staffNames(j) & " on weekend or holiday: " & currentDate
                    GoTo NextAdmin
                End If
            End If

            ' 管理者に日勤を割り当てる
            If staffAttributes(j) = "4" Then
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If hoursWorked(j) + 8.5 <= staffMaxHours(j) Then
                        Call AssignShift(wsOutput, j, i, g_marks.dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, "day", _
                            staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                        Exit For ' 一人の管理者が割り当てられたら次の日付へ
                    End If
                End If
            End If
NextAdmin:
        Next j
    Next i

' ステップ2: 夜勤リーダーを割り付ける
Print #1, "Step 11: Assigning night shift leaders"
For i = 1 To shiftDates.count
    currentDate = shiftDates(i)
    nightShiftLeaderAssigned = 0
    nightShiftLeaderRequired = nightShiftLeaderMin

    Do While nightShiftLeaderAssigned < nightShiftLeaderRequired
        Dim leaderAssigned As Boolean
        leaderAssigned = False

        ' シャッフルされたスタッフリストを作成
        Dim shuffledStaff As New Collection
        For k = 1 To UBound(staffNames)
            shuffledStaff.Add k
        Next k
        ShuffleCollection shuffledStaff ' スタッフリストをシャッフル

        ' 条件を満たすシャッフルされたスタッフを探す
        For Each staffIndex In shuffledStaff
            j = staffIndex ' スタッフのインデックス
            If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                If (staffAttributes(j) = "2" Or staffAttributes(j) = "3" Or staffAttributes(j) = "4") Then
                    If hoursWorked(j) + 14.5 <= staffMaxHours(j) And nightShiftCountWorked(j) < staffNightShiftMax(j) Then
                        ' 新しい属性を確認する
                        If IsAvailableForNightShift(j, currentDate) Then
                            If CanAssign(wsOutput, j, i, "night", True, staffNames, shiftDates) Then
                                If Not HasIncompatibleStaff(wsOutput, g_staffCompatibility, j, i, "night", g_marks.dayShiftMark, g_marks.nightShiftMark) Then
                                    Call AssignShift(wsOutput, j, i, g_marks.nightShiftMark, 14.5, hoursWorked, staffShifts, staffNames, _
                                        "night", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                                    nightShiftLeaderAssigned = nightShiftLeaderAssigned + 1
                                    leaderAssigned = True
                                    Exit For ' リーダーが割り当てられたので次のスタッフを探す
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next staffIndex

        ' スタッフが見つからなかった場合、制約を緩和して再度探す
        If Not leaderAssigned Then
            Print #1, "No suitable night shift leader found for date " & currentDate & ". Relaxing constraints."
            For j = 1 To UBound(staffNames)
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If (staffAttributes(j) = "2" Or staffAttributes(j) = "3" Or staffAttributes(j) = "4") Then
                        ' 最大勤務時間の制限を緩和
                        If hoursWorked(j) + 14.5 <= staffMaxHours(j) + 14.5 Then
                            ' 夜勤の最大回数の制限を緩和
                            If nightShiftCountWorked(j) < staffNightShiftMax(j) + 1 Then
                                ' 新しい属性を確認する
                                If IsAvailableForNightShift(j, currentDate) Then
                                    ' CanAssign関数のstrictCheckをFalseに設定
                                    If CanAssign(wsOutput, j, i, "night", False, staffNames, shiftDates) Then
                                        ' 相性のチェックを無視
                                        ' 割り当てる
                                        Call AssignShift(wsOutput, j, i, g_marks.nightShiftMark, 14.5, hoursWorked, staffShifts, staffNames, _
                                            "night", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                                        nightShiftLeaderAssigned = nightShiftLeaderAssigned + 1
                                        leaderAssigned = True
                                        Print #1, "Assigned night shift leader with relaxed constraints to " & staffNames(j)
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        End If

        ' それでもスタッフが見つからなかった場合、警告を出して強制的に割り当てる
        If Not leaderAssigned Then
            Print #1, "Warning: No night shift leader could be assigned for date " & currentDate & ". Assigning the first available leader."
            For j = 1 To UBound(staffNames)
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If (staffAttributes(j) = "2" Or staffAttributes(j) = "3" Or staffAttributes(j) = "4") Then
                        ' 新しい属性を確認する
                        If IsAvailableForNightShift(j, currentDate) Then
                            ' 制限を全て無視して割り当てる
                            Call AssignShift(wsOutput, j, i, g_marks.nightShiftMark, 14.5, hoursWorked, staffShifts, staffNames, _
                                "night", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                            nightShiftLeaderAssigned = nightShiftLeaderAssigned + 1
                            leaderAssigned = True
                            Print #1, "Forcefully assigned night shift leader to " & staffNames(j)
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If

        ' 最終的にリーダーが割り当てられなかった場合、ループを抜ける
        If Not leaderAssigned Then
            Print #1, "Error: Unable to assign any night shift leader for date " & currentDate
            Exit Do
        End If

    Loop ' Do Whileの終了

Next i


' ステップ3: 日勤リーダーを割り付ける
Print #1, "Step 12: Assigning day shift leaders"
For i = 1 To shiftDates.count
    currentDate = shiftDates(i)
    dayShiftLeaderAssigned = 0
    dayShiftLeaderRequired = dayShiftLeaderMin

    Do While dayShiftLeaderAssigned < dayShiftLeaderRequired
        
        leaderAssigned = False

        ' リーダー候補のスタッフを探す
        For j = 1 To UBound(staffNames)
            If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                If (staffAttributes(j) = "2" Or staffAttributes(j) = "3" Or staffAttributes(j) = "4") Then
                    If hoursWorked(j) + 8.5 <= staffMaxHours(j) Then
                        If CanAssign(wsOutput, j, i, "day", True, staffNames, shiftDates) Then
                            If Not HasIncompatibleStaff(wsOutput, g_staffCompatibility, j, i, "day", g_marks.dayShiftMark, g_marks.nightShiftMark) Then
                                Call AssignShift(wsOutput, j, i, g_marks.dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, _
                                    "day", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                                dayShiftLeaderAssigned = dayShiftLeaderAssigned + 1
                                leaderAssigned = True
                                Exit For ' リーダーが割り当てられたので内側のループを抜ける
                            End If
                        End If
                    End If
                End If
            End If
        Next j

        ' スタッフが見つからなかった場合、制限を緩和して再度探す
        If Not leaderAssigned Then
            Print #1, "No suitable day shift leader found for date " & currentDate & ". Relaxing constraints."
            For j = 1 To UBound(staffNames)
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If (staffAttributes(j) = "2" Or staffAttributes(j) = "3" Or staffAttributes(j) = "4") Then
                        ' 最大勤務時間の制限を緩和
                        If hoursWorked(j) + 8.5 <= staffMaxHours(j) + 8.5 Then
                            ' CanAssign関数のstrictCheckをFalseに設定
                            If CanAssign(wsOutput, j, i, "day", False, staffNames, shiftDates) Then
                                ' 相性のチェックを無視
                                ' 割り当てる
                                Call AssignShift(wsOutput, j, i, g_marks.dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, _
                                    "day", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                                dayShiftLeaderAssigned = dayShiftLeaderAssigned + 1
                                leaderAssigned = True
                                Print #1, "Assigned day shift leader with relaxed constraints to " & staffNames(j)
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next j
        End If

        ' それでもスタッフが見つからなかった場合、警告を出して強制的に割り当てる
        If Not leaderAssigned Then
            Print #1, "Warning: No day shift leader could be assigned for date " & currentDate & ". Assigning the first available leader."
            For j = 1 To UBound(staffNames)
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If (staffAttributes(j) = "2" Or staffAttributes(j) = "3" Or staffAttributes(j) = "4") Then
                        ' 制限を全て無視して割り当てる
                        Call AssignShift(wsOutput, j, i, g_marks.dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, _
                            "day", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                        dayShiftLeaderAssigned = dayShiftLeaderAssigned + 1
                        leaderAssigned = True
                        Print #1, "Forcefully assigned day shift leader to " & staffNames(j)
                        Exit For
                    End If
                End If
            Next j
        End If

        ' 最終的にリーダーが割り当てられなかった場合、ループを抜ける
        If Not leaderAssigned Then
            Print #1, "Error: Unable to assign any day shift leader for date " & currentDate
            Exit Do
        End If

    Loop ' Do Whileの終了

Next i


' ステップ4: 夜勤の最低限勤務を割り付ける
Print #1, "Step 13: Assigning minimum night shifts"

For i = 1 To shiftDates.count
    currentDate = shiftDates(i)

    ' スタッフリストをシャッフル
    Set availableStaff = New Collection
    For j = 1 To UBound(staffNames)
        availableStaff.Add staffNames(j)
    Next j
    ShuffleCollection availableStaff

    dayOfWeek = Weekday(currentDate, vbSunday)
    If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
        nightShiftCount = weekdayNightShiftMin
    Else
        nightShiftCount = weekendNightShiftMin
    End If

    ' 公休日の場合は週末の夜勤人数を適用
    For Each holidayDate In holidayDates
        If currentDate = holidayDate Then
            nightShiftCount = weekendNightShiftMin
            Exit For
        End If
    Next holidayDate

    ' 現在の夜勤者数を確認
    nightShiftAssigned = nightShiftDailyCount(i)

    ' 夜勤の最低人数に達していない場合、割り当てを行う
    If nightShiftAssigned < nightShiftCount Then
        Set availableStaff = New Collection
        Set highPriorityStaff = New Collection ' 日勤上限0のスタッフのリスト

        For j = 1 To UBound(staffNames)
            ' スタッフの現在の夜勤数と最大夜勤日数の差を計算
            Dim remainingNightShifts As Integer
            remainingNightShifts = staffNightShiftMax(j) - nightShiftCountWorked(j)
            ' 夜勤の上限に達していないスタッフを収集
            If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                If remainingNightShifts > 0 And hoursWorked(j) + 14.5 <= staffMaxHours(j) Then
                    ' 新しい属性を確認する
                    If IsAvailableForNightShift(j, currentDate) Then
                        ' 日勤上限が '0' のスタッフは別リストに追加
                        If staffDayShiftMax(j) = 0 Then
                            highPriorityStaff.Add j ' 最優先で追加
                        Else
                            availableStaff.Add j ' 通常のスタッフをリストに追加
                        End If
                    End If
                End If
            End If
        Next j

        ' 日勤上限が0のスタッフを先に処理
        If highPriorityStaff.count > 0 Then
            For Each staffIndex In highPriorityStaff
                If nightShiftAssigned < nightShiftCount Then
                    ' **CanAssign関数を呼び出してチェック**
                    If CanAssign(wsOutput, staffIndex, i, "night", True, staffNames, shiftDates) Then
                        ' 相性の確認
                        If Not HasIncompatibleStaff(wsOutput, g_staffCompatibility, staffIndex, i, "night", g_marks.dayShiftMark, g_marks.nightShiftMark) Then
                            Call AssignShift(wsOutput, staffIndex, i, g_marks.nightShiftMark, 14.5, hoursWorked, staffShifts, staffNames, _
                                "night", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                            nightShiftAssigned = nightShiftAssigned + 1
                        End If
                    End If
                Else
                    Exit For
                End If
            Next staffIndex
        End If

        ' 通常のスタッフを処理
        If availableStaff.count > 0 Then
            ' スタッフを残りの夜勤可能回数でソート
            Call SortStaffByRemainingNightShifts(availableStaff, nightShiftCountWorked, staffNightShiftMax)
        End If

        For Each staffIndex In availableStaff
            If nightShiftAssigned < nightShiftCount Then
                ' **CanAssign関数を呼び出してチェック**
                If CanAssign(wsOutput, staffIndex, i, "night", True, staffNames, shiftDates) Then
                    ' 相性の確認
                    If Not HasIncompatibleStaff(wsOutput, g_staffCompatibility, staffIndex, i, "night", g_marks.dayShiftMark, g_marks.nightShiftMark) Then
                        Call AssignShift(wsOutput, staffIndex, i, g_marks.nightShiftMark, 14.5, hoursWorked, staffShifts, staffNames, _
                            "night", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                        nightShiftAssigned = nightShiftAssigned + 1
                    End If
                End If
            Else
                Exit For
            End If
        Next staffIndex
    End If
Next i


    ' ステップ5: 日勤の最低限勤務を割り付ける
    Print #1, "Step 14: Assigning minimum day shifts"
    For i = 1 To shiftDates.count
        currentDate = shiftDates(i)

        ' スタッフリストをシャッフル
        Set availableStaff = New Collection
        For j = 1 To UBound(staffNames)
            availableStaff.Add staffNames(j)
        Next j
        ShuffleCollection availableStaff

        dayOfWeek = Weekday(currentDate, vbSunday)
        If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
            dayShiftCount = weekdayDayShiftMin
        Else
            dayShiftCount = weekendDayShiftMin
        End If

        For Each holidayDate In holidayDates
            If currentDate = holidayDate Then
                dayShiftCount = weekendDayShiftMin
                Exit For
            End If
        Next holidayDate

        ' 現在の日勤者数を確認
        dayShiftAssigned = dayShiftDailyCount(i)

        ' 日勤の最低人数に達していない場合、割り当てを行う
        If dayShiftAssigned < dayShiftCount Then
            Set availableStaff = New Collection
            For j = 1 To UBound(staffNames)
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If hoursWorked(j) + 8.5 <= staffMaxHours(j) Then
                        If CanAssign(wsOutput, j, i, "day", True, staffNames, shiftDates) Then
                            availableStaff.Add j
                        End If
                    End If
                End If
            Next j

            ' 勤務時間が少ない順にソート
            If availableStaff.count > 0 Then
                ' ランダムにシャッフル
                ShuffleCollection availableStaff
                Call SortStaffByHoursWorked(availableStaff, hoursWorked, totalShiftsAssigned)
            End If

            For Each staffIndex In availableStaff
                If dayShiftAssigned < dayShiftCount Then
                    If Not HasIncompatibleStaff(wsOutput, g_staffCompatibility, staffIndex, i, "day", g_marks.dayShiftMark, g_marks.nightShiftMark) Then
                        Call AssignShift(wsOutput, staffIndex, i, g_marks.dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, _
                            "day", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                        dayShiftAssigned = dayShiftAssigned + 1
                    End If
                Else
                    Exit For
                End If
            Next staffIndex
        End If
    Next i

    ' ステップ6: 夜勤の余剰勤務を割り付ける
    Print #1, "Step 15: Assigning extra night shifts"
    For i = 1 To shiftDates.count
        currentDate = shiftDates(i)

        dayOfWeek = Weekday(currentDate, vbSunday)
        If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
            nightShiftCount = weekdayNightShiftMax
        Else
            nightShiftCount = weekendNightShiftMax
        End If

        For Each holidayDate In holidayDates
            If currentDate = holidayDate Then
                nightShiftCount = weekendNightShiftMax
                Exit For
            End If
        Next holidayDate

        ' 既に割り当てられている夜勤者数を確認
        nightShiftAssigned = nightShiftDailyCount(i)

        ' 夜勤の最大人数に達していない場合、余剰勤務を割り付ける
        If nightShiftAssigned < nightShiftCount Then
            Set availableStaff = New Collection
            For j = 1 To UBound(staffNames)
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If hoursWorked(j) + 14.5 <= staffMaxHours(j) And nightShiftCountWorked(j) < staffNightShiftMax(j) Then
                        ' 新しい属性を確認する
                        If IsAvailableForNightShift(j, currentDate) Then
                            If CanAssign(wsOutput, j, i, "night", True, staffNames, shiftDates) Then
                                availableStaff.Add j
                            End If
                        End If
                    End If
                End If
            Next j

            ' 勤務時間が少ない順にソート
            If availableStaff.count > 0 Then
                ' ランダムにシャッフル
                ShuffleCollection availableStaff
                Call SortStaffByHoursWorked(availableStaff, hoursWorked, totalShiftsAssigned)
            End If

            For Each staffIndex In availableStaff
                If nightShiftAssigned < nightShiftCount Then
                    If Not HasIncompatibleStaff(wsOutput, g_staffCompatibility, staffIndex, i, "night", g_marks.dayShiftMark, g_marks.nightShiftMark) Then
                        Call AssignShift(wsOutput, staffIndex, i, g_marks.nightShiftMark, 14.5, hoursWorked, staffShifts, staffNames, _
                            "night", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                        nightShiftAssigned = nightShiftAssigned + 1
                    End If
                Else
                    Exit For
                End If
            Next staffIndex
        End If
    Next i

    ' ステップ7: 日勤の余剰勤務を割り付ける
    Print #1, "Step 16: Assigning extra day shifts"
    For i = 1 To shiftDates.count
        currentDate = shiftDates(i)

        dayOfWeek = Weekday(currentDate, vbSunday)
        If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
            dayShiftCount = weekdayDayShiftMax
        Else
            dayShiftCount = weekendDayShiftMax
        End If

        For Each holidayDate In holidayDates
            If currentDate = holidayDate Then
                dayShiftCount = weekendDayShiftMax
                Exit For
            End If
        Next holidayDate

        ' 既に割り当てられている日勤者数を確認
        dayShiftAssigned = dayShiftDailyCount(i)

        ' 日勤の最大人数に達していない場合、余剰勤務を割り付ける
        If dayShiftAssigned < dayShiftCount Then
            Set availableStaff = New Collection
            For j = 1 To UBound(staffNames)
                If wsOutput.Cells(j + 2, i + 2).Value = "" Then
                    If hoursWorked(j) + 8.5 <= staffMaxHours(j) Then
                        If CanAssign(wsOutput, j, i, "day", True, staffNames, shiftDates) Then
                            availableStaff.Add j
                        End If
                    End If
                End If
            Next j

            ' 勤務時間が少ない順にソート
            If availableStaff.count > 0 Then
                ' ランダムにシャッフル
                ShuffleCollection availableStaff
                Call SortStaffByHoursWorked(availableStaff, hoursWorked, totalShiftsAssigned)
            End If

            For Each staffIndex In availableStaff
                If dayShiftAssigned < dayShiftCount Then
                    If Not HasIncompatibleStaff(wsOutput, g_staffCompatibility, staffIndex, i, "day", g_marks.dayShiftMark, g_marks.nightShiftMark) Then
                        Call AssignShift(wsOutput, staffIndex, i, g_marks.dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, _
                            "day", staffAttributes, currentDate, nightShiftDailyCount, nightShiftCountWorked, dayShiftDailyCount, shiftDates)
                        dayShiftAssigned = dayShiftAssigned + 1
                    End If
                Else
                    Exit For
                End If
            Next staffIndex
        End If
    Next i

    On Error GoTo ErrorHandler

' 各スタッフの勤務時間を最終日に記載
Print #1, "Step 17: Recording hours worked"
On Error GoTo HoursRecordingError
For i = 1 To UBound(staffNames)
    If staffNames(i) <> "" Then
        wsOutput.Cells(i + 2, shiftDates.count + 3).Value = hoursWorked(i)
    End If
Next i

On Error GoTo ErrorHandler

' 日勤数、夜勤数、総勤務時間数の計算式を追加
For i = 1 To UBound(staffNames)
    If staffNames(i) <> "" Then
        wsOutput.Cells(i + 2, shiftDates.count + 4).Formula = "=COUNTIF(C" & i + 2 & ":" & wsOutput.Cells(i + 2, shiftDates.count + 2).Address(False, False) & ", """ & g_marks.dayShiftMark & """)"
        wsOutput.Cells(i + 2, shiftDates.count + 5).Formula = "=COUNTIF(C" & i + 2 & ":" & wsOutput.Cells(i + 2, shiftDates.count + 2).Address(False, False) & ", """ & g_marks.nightShiftMark & """)"
        wsOutput.Cells(i + 2, shiftDates.count + 6).Formula = "=" & wsOutput.Cells(i + 2, shiftDates.count + 4).Address(False, False) & "*8.5 + " & wsOutput.Cells(i + 2, shiftDates.count + 5).Address(False, False) & "*14.5"
    End If
Next i

' シフト表の1行目と1列目を固定
wsOutput.Cells(3, 3).Select
ActiveWindow.FreezePanes = True

' 見出しを追加
wsOutput.Cells(1, shiftDates.count + 4).Value = "日勤数"
wsOutput.Cells(1, shiftDates.count + 5).Value = "夜勤数"
wsOutput.Cells(1, shiftDates.count + 6).Value = "総勤務時間"

' 日毎のシフトカウントを修正
For i = 1 To shiftDates.count
    ' 日勤のカウント
    wsOutput.Cells(staffCount + 3, i + 2).Formula = "=COUNTIF(" & wsOutput.Cells(3, i + 2).Address(False, False) & ":" & _
        wsOutput.Cells(staffCount + 2, i + 2).Address(False, False) & ", """ & g_marks.dayShiftMark & """)"
    ' 夜勤のカウント
    wsOutput.Cells(staffCount + 4, i + 2).Formula = "=COUNTIF(" & wsOutput.Cells(3, i + 2).Address(False, False) & ":" & _
        wsOutput.Cells(staffCount + 2, i + 2).Address(False, False) & ", """ & g_marks.nightShiftMark & """)"
    ' 休みのカウント
    wsOutput.Cells(staffCount + 5, i + 2).Formula = "=COUNTIF(" & wsOutput.Cells(3, i + 2).Address(False, False) & ":" & _
        wsOutput.Cells(staffCount + 2, i + 2).Address(False, False) & ", """ & g_marks.holidayMark & """)"
    ' 日勤リーダーのカウント
    wsOutput.Cells(staffCount + 6, i + 2).Formula = "=SUMPRODUCT(--(" & wsOutput.Cells(3, i + 2).Address(False, False) & ":" & _
        wsOutput.Cells(staffCount + 2, i + 2).Address(False, False) & "=""" & g_marks.dayShiftMark & """),--((" & _
        wsOutput.Cells(3, 2).Address(False, False) & ":" & wsOutput.Cells(staffCount + 2, 2).Address(False, False) & "=2) + (" & _
        wsOutput.Cells(3, 2).Address(False, False) & ":" & wsOutput.Cells(staffCount + 2, 2).Address(False, False) & "=3) + (" & _
        wsOutput.Cells(3, 2).Address(False, False) & ":" & wsOutput.Cells(staffCount + 2, 2).Address(False, False) & "=4)))"
    ' 夜勤リーダーのカウント
    wsOutput.Cells(staffCount + 7, i + 2).Formula = "=SUMPRODUCT(--(" & wsOutput.Cells(3, i + 2).Address(False, False) & ":" & _
        wsOutput.Cells(staffCount + 2, i + 2).Address(False, False) & "=""" & g_marks.nightShiftMark & """),--((" & _
        wsOutput.Cells(3, 2).Address(False, False) & ":" & wsOutput.Cells(staffCount + 2, 2).Address(False, False) & "=2) + (" & _
        wsOutput.Cells(3, 2).Address(False, False) & ":" & wsOutput.Cells(staffCount + 2, 2).Address(False, False) & "=3) + (" & _
        wsOutput.Cells(3, 2).Address(False, False) & ":" & wsOutput.Cells(staffCount + 2, 2).Address(False, False) & "=4)))"
Next i
    ' カウント行のラベルを追加
    wsOutput.Cells(staffCount + 3, 1).Value = "日勤者数"
    wsOutput.Cells(staffCount + 4, 1).Value = "夜勤者数"
    wsOutput.Cells(staffCount + 5, 1).Value = "休み者数"
    wsOutput.Cells(staffCount + 6, 1).Value = "日勤リーダー可能スタッフ数"
    wsOutput.Cells(staffCount + 7, 1).Value = "夜勤リーダー可能スタッフ数"
' 空欄セルを休み("X")で埋める
Print #1, "Step 18: Filling empty cells with 'X'"
For i = 1 To UBound(staffNames)
    For j = 1 To shiftDates.count
        If wsOutput.Cells(i + 2, j + 2).Value = "" Then
            ' 直前が夜勤でかつ連続夜勤適性がある場合、休みを割り当てない
            If g_staffCompatibility(i) = "1" And wsOutput.Cells(i + 2, j + 1).Value = g_marks.nightShiftMark Then
                Print #1, "Skipping rest assignment after consecutive night shift for " & staffNames(i) & " on " & shiftDates(j)
                GoTo NextDay
            End If
            ' 通常の休みの割り当て
            wsOutput.Cells(i + 2, j + 2).Value = g_marks.holidayMark
            Print #1, "Assigned rest to " & staffNames(i) & " on " & shiftDates(j)
        End If
NextDay:
    Next j
Next i

    On Error GoTo ErrorHandler

    Print #1, "Shift schedule created successfully"
    Close #1
    MsgBox "Shift schedule created successfully"
    Exit Sub

    ' エラーハンドラを追加
DataError:
    Print #1, "Data Error: " & Err.Description & " at line " & Erl & " (Step 2)"
    Resume Next

StaffDataError:
    Print #1, "Staff Data Error: " & Err.Description & " at line " & Erl & " (Step 3)"
    Resume Next

HolidayDataError:
    Print #1, "Holiday Data Error: " & Err.Description & " at line " & Erl & " (Step 4)"
    Resume Next

OutputSheetError:
    Print #1, "Output Sheet Error: " & Err.Description & " at line " & Erl & " (Step 5)"
    Resume Next

DateRowError:
    Print #1, "Date Row Error: " & Err.Description & " at line " & Erl & " (Step 7)"
    Resume Next

StaffNameColumnError:
    Print #1, "Staff Name Column Error: " & Err.Description & " at line " & Erl & " (Step 8)"
    Resume Next

HolidaySettingError:
    Print #1, "Holiday Setting Error: " & Err.Description & " at line " & Erl & " (Step 9)"
    Resume Next

HoursRecordingError:
    Print #1, "Hours Recording Error: " & Err.Description & " at line " & Erl & " (Step 17)"
    Resume Next

ErrorHandler:
    Print #1, "Error: " & Err.Description & " at line " & Erl
    Close #1
    MsgBox "An error occurred: " & Err.Description
End Sub
Sub AssignShift(ws As Worksheet, ByVal staffIndex As Integer, ByVal dateIndex As Integer, ByVal shiftMark As String, ByVal hours As Double, _
    ByRef hoursWorked() As Double, ByRef staffShifts() As Collection, ByRef staffNames() As String, ByVal shiftType As String, _
    ByRef staffAttributes() As String, ByVal currentDate As Date, ByRef nightShiftDailyCount() As Integer, _
    ByRef nightShiftCountWorked() As Integer, ByRef dayShiftDailyCount() As Integer, ByRef shiftDates As Collection)

    ' 日勤シフトの場合は割り当て前に必ず上限チェックを行う
    If shiftType = "day" Then
        ' 日勤シフトが上限を超えていないかチェック
        If dayShiftDailyCount(dateIndex) >= staffDayShiftMax(staffIndex) Then
            ' 上限超過時の処理（ログなど）
            Print #1, "Cannot assign day shift to " & staffNames(staffIndex) & " due to exceeding day shift max limit."
            Exit Sub
        End If
    End If

    ' 総シフト数をインクリメント
    totalShiftsAssigned(staffIndex) = totalShiftsAssigned(staffIndex) + 1

    Dim canAssignShift As Boolean
    canAssignShift = False

    ' 夜勤の場合の処理
    If shiftType = "night" Then
        ' 連続夜勤適性のあるスタッフの場合
        If g_consecutiveNightShiftAbility(staffIndex) = "1" Then
            ' まず連続夜勤の割り当てを試みる
            If CanAssign(ws, staffIndex, dateIndex, "night", True, staffNames, shiftDates, 4) Then
                ' 連続夜勤の割り当て処理
                ' 月末の2日間の場合、連続夜勤を割り当てない
                Dim lastDayOfMonth As Date
                lastDayOfMonth = DateSerial(Year(currentDate), Month(currentDate) + 1, 0)
                If currentDate >= lastDayOfMonth - 1 Then
                    Print #1, "Current date is within the last two days of the month. Not assigning consecutive night shifts."
                Else
                    ' 連続夜勤を割り当てても勤務時間の上限を超えないかをチェック
                    Dim totalHoursForConsecutiveNightShifts As Double
                    totalHoursForConsecutiveNightShifts = hours * 2 ' 連続夜勤2日分の時間
                    If hoursWorked(staffIndex) + totalHoursForConsecutiveNightShifts <= staffMaxHours(staffIndex) Then
                        ' 連続夜勤の割り当て処理をここに記述
                        ' 1日目の夜勤割り当て
                        ws.Cells(staffIndex + 2, dateIndex + 2).Value = shiftMark
                        hoursWorked(staffIndex) = hoursWorked(staffIndex) + hours
                        staffShifts(staffIndex).Add shiftType
                        nightShiftCountWorked(staffIndex) = nightShiftCountWorked(staffIndex) + 1
                        nightShiftDailyCount(dateIndex) = nightShiftDailyCount(dateIndex) + 1

                        ' 1日目の夜勤明けマークを翌日に割り当てる
                        If dateIndex + 1 <= g_totalDates Then
                            ws.Cells(staffIndex + 2, dateIndex + 3).Value = g_marks.nightShiftAfterMark
                        End If

                        ' 2日目の夜勤割り当て
                        If CanAssign(ws, staffIndex, dateIndex + 2, "night", False, staffNames, shiftDates) Then
                            ' 2日目の夜勤を割り当て
                            ws.Cells(staffIndex + 2, dateIndex + 4).Value = shiftMark
                            hoursWorked(staffIndex) = hoursWorked(staffIndex) + hours
                            staffShifts(staffIndex).Add shiftType
                            nightShiftCountWorked(staffIndex) = nightShiftCountWorked(staffIndex) + 1
                            nightShiftDailyCount(dateIndex + 2) = nightShiftDailyCount(dateIndex + 2) + 1

                            ' 2日目の夜勤明けマークを翌日に割り当てる
                            If dateIndex + 3 <= g_totalDates Then
                                ws.Cells(staffIndex + 2, dateIndex + 5).Value = g_marks.nightShiftAfterMark
                            End If

                            ' 休みマークをその次の日に割り当てる
                            If dateIndex + 4 <= g_totalDates Then
                                ws.Cells(staffIndex + 2, dateIndex + 6).Value = g_marks.holidayMark
                            End If

                            ' 割り当て完了
                            Exit Sub
                        Else
                            Print #1, "Cannot assign second night shift for consecutive night shifts to " & staffNames(staffIndex) & "."
                        End If
                    Else
                        Print #1, "Total hours after assigning consecutive night shifts would exceed the maximum allowed hours. Not assigning consecutive night shifts."
                    End If
                End If
            End If
        End If

        ' 単発夜勤の割り当てを試みる
        If CanAssign(ws, staffIndex, dateIndex, "night", True, staffNames, shiftDates, 2) Then
            canAssignShift = True
        Else
            ' 夜勤を割り当てられない場合はサブルーチンを終了
            Exit Sub
        End If
    Else
        ' 日勤や他のシフトの場合の処理
        If CanAssign(ws, staffIndex, dateIndex, shiftType, True, staffNames, shiftDates, 2) Then
            canAssignShift = True
        Else
            ' シフトを割り当てられない場合はサブルーチンを終了
            Exit Sub
        End If
    End If

    ' シフトの割り当てを行う
    If canAssignShift Then
        ' 勤務時間の上限チェック
        If hoursWorked(staffIndex) + hours <= staffMaxHours(staffIndex) Then
            ' シフトを割り当てる
            ws.Cells(staffIndex + 2, dateIndex + 2).Value = shiftMark
            hoursWorked(staffIndex) = hoursWorked(staffIndex) + hours
            staffShifts(staffIndex).Add shiftType
            Print #1, currentDate & ": Assigned " & shiftType & " shift to: " & staffNames(staffIndex) & ", Total Hours Worked: " & hoursWorked(staffIndex)

            ' 日毎のカウントを更新
            If shiftType = "night" Then
                nightShiftCountWorked(staffIndex) = nightShiftCountWorked(staffIndex) + 1
                nightShiftDailyCount(dateIndex) = nightShiftDailyCount(dateIndex) + 1

                ' 翌日の夜勤明けマークを割り当てる
                If dateIndex + 1 <= UBound(nightShiftDailyCount) Then
                    ws.Cells(staffIndex + 2, dateIndex + 2 + 1).Value = g_marks.nightShiftAfterMark
                End If
            ElseIf shiftType = "day" Then
                dayShiftDailyCount(dateIndex) = dayShiftDailyCount(dateIndex) + 1
            End If
        Else
            Print #1, "Cannot assign " & shiftType & " shift to " & staffNames(staffIndex) & " because it would exceed maximum allowed hours."
            Exit Sub
        End If
    End If

    ' リーダーのカウントを更新
    If shiftType = "day" And (staffAttributes(staffIndex) = "2" Or staffAttributes(staffIndex) = "3" Or staffAttributes(staffIndex) = "4") Then
        ws.Cells(UBound(staffNames) + 6, dateIndex + 2).Value = ws.Cells(UBound(staffNames) + 6, dateIndex + 2).Value + 1
    ElseIf shiftType = "night" And (staffAttributes(staffIndex) = "2" Or staffAttributes(staffIndex) = "3" Or staffAttributes(staffIndex) = "4") Then
        ws.Cells(UBound(staffNames) + 7, dateIndex + 2).Value = ws.Cells(UBound(staffNames) + 7, dateIndex + 2).Value + 1
    End If
End Sub

Function CanAssign(ws As Worksheet, ByVal staffIndex As Integer, ByVal dateIndex As Integer, ByVal shiftType As String, ByVal strictCheck As Boolean, ByRef staffNames() As String, ByRef shiftDates As Collection, Optional ByVal daysToCheck As Integer = 2) As Boolean
    Dim i As Integer
    Dim maxContinuousShifts As Integer
    Dim continuousShifts As Integer
    continuousShifts = 0
    CanAssign = True

    ' スタッフごとに最大連続日勤日数を設定
    If staffMaxConsecutiveDayShifts(staffIndex) >= 1 Then
        maxContinuousShifts = staffMaxConsecutiveDayShifts(staffIndex)
    Else
        maxContinuousShifts = 5 ' デフォルト値（必要に応じて変更）
    End If

    ' 夜勤を割り振る前に指定された日数後までのシフト確認
    If shiftType = "night" Then
        If Not CheckFutureDays(ws, staffIndex, dateIndex, daysToCheck, g_marks.holidayMark, g_marks.nightShiftMark, g_marks.nightShiftAfterMark, shiftDates) Then
            CanAssign = False
            Exit Function
        End If
    End If

    ' 日勤を割り当てる際のチェックを追加
    If shiftType = "day" Then
        ' 前日が夜勤または夜勤明けマークの場合
        If dateIndex > 1 Then
            Dim previousShift As String
            previousShift = ws.Cells(staffIndex + 2, dateIndex + 2 - 1).Value ' 前日のセルを参照
            If previousShift = g_marks.nightShiftMark Or previousShift = g_marks.nightShiftAfterMark Then
                ' 連続夜勤適性がない場合は日勤を割り当てない
                If g_consecutiveNightShiftAbility(staffIndex) <> "1" Then
                    CanAssign = False
                    Exit Function
                End If
            End If
        End If
    End If

    ' 連続日勤のチェック
    If shiftType = "day" Then
        continuousShifts = 1 ' 現在の日勤を含む

        ' 過去の日勤をカウント
        For i = dateIndex - 1 To dateIndex - maxContinuousShifts Step -1
            If i < 1 Then
                ' 月初で、前月末の連続日勤数を考慮
                If previousMonthConsecutiveDayShift(staffIndex) > 0 Then
                    continuousShifts = continuousShifts + previousMonthConsecutiveDayShift(staffIndex)
                End If
                Exit For
            End If
            If ws.Cells(staffIndex + 2, i + 2).Value = g_marks.dayShiftMark Then
                continuousShifts = continuousShifts + 1
            Else
                Exit For
            End If
        Next i

        ' 未来の日勤をカウント
        For i = dateIndex + 1 To dateIndex + maxContinuousShifts
            If i > g_totalDates Then Exit For
            If ws.Cells(staffIndex + 2, i + 2).Value = g_marks.dayShiftMark Then
                continuousShifts = continuousShifts + 1
            Else
                Exit For
            End If
        Next i

        ' 連続日勤が最大値を超える場合、シフトを割り当てない
        If continuousShifts > maxContinuousShifts Then
            CanAssign = False
            Exit Function
        End If
    End If
    ' 連続夜勤のチェック
    If shiftType = "night" Then
        continuousShifts = 0 ' 指定日の前後を含む連続夜勤チェック

        ' 過去の夜勤をカウント
        For i = dateIndex To dateIndex - 5 Step -1
            If i < 1 Then Exit For
            If ws.Cells(staffIndex + 2, i + 2).Value = g_marks.nightShiftMark Then
                continuousShifts = continuousShifts + 1
            ElseIf ws.Cells(staffIndex + 2, i + 2).Value <> "" Then
                Exit For
            End If
        Next i

        ' 未来の夜勤をカウント
        For i = dateIndex + 1 To dateIndex + 5
            If i > g_totalDates Then Exit For
            If ws.Cells(staffIndex + 2, i + 2).Value = g_marks.nightShiftMark Then
                continuousShifts = continuousShifts + 1
            ElseIf ws.Cells(staffIndex + 2, i + 2).Value <> "" Then
                Exit For
            End If
        Next i

        If continuousShifts >= maxContinuousShifts Then
            CanAssign = False
            Exit Function
        End If
    End If

    ' 夜勤明けの次の日に夜勤を割り当てる場合のチェック
    If shiftType = "night" And strictCheck Then
        If dateIndex > 1 And ws.Cells(staffIndex + 2, dateIndex + 2 - 1).Value = g_marks.nightShiftAfterMark Then
            ' 連続夜勤適性がある場合は次の夜勤を優先
            If g_consecutiveNightShiftAbility(staffIndex) = "1" Then
                CanAssign = True
            ElseIf ws.Cells(staffIndex + 2, dateIndex).Value = g_marks.nightShiftMark Then
                CanAssign = False
                Exit Function
            End If
        End If
    End If
End Function


Function HasIncompatibleStaff(ws As Worksheet, staffCompatibility() As String, ByVal staffIndex As Integer, ByVal dateIndex As Integer, ByVal shiftType As String, ByVal dayShiftMark As String, ByVal nightShiftMark As String) As Boolean
    Dim i As Integer
    Dim incompatibility As String
    incompatibility = staffCompatibility(staffIndex)
    If incompatibility = "" Then
        HasIncompatibleStaff = False
        Exit Function
    End If

    For i = 1 To UBound(staffCompatibility)
        If i <> staffIndex And staffCompatibility(i) = incompatibility Then
            If shiftType = "day" And ws.Cells(i + 2, dateIndex + 2).Value = dayShiftMark Then
                HasIncompatibleStaff = True
                Exit Function
            ElseIf shiftType = "night" And ws.Cells(i + 2, dateIndex + 2).Value = nightShiftMark Then
                HasIncompatibleStaff = True
                Exit Function
            End If
        End If
    Next i
    HasIncompatibleStaff = False
End Function

Function CheckFutureDays(ws As Worksheet, ByVal staffIndex As Integer, ByVal dateIndex As Integer, _
    ByVal daysToCheck As Integer, ByVal holidayMark As String, ByVal nightShiftMark As String, _
    ByVal nightShiftAfterMark As String, ByRef shiftDates As Collection) As Boolean

    Dim i As Integer
    Dim cellValue As String
    CheckFutureDays = True ' 初期化

    For i = 1 To daysToCheck
        If dateIndex + i > g_totalDates Then
            ' 月末を超える場合は何もしない
        Else
            ' スタッフの指定休日を確認
            Dim futureDate As Date
            futureDate = shiftDates(dateIndex + i)
            If IsStaffHoliday(staffIndex, futureDate) Then
                ' 指定休日がある場合、割り当て不可
                CheckFutureDays = False
                Exit Function
            End If

            ' 未来の日に他のシフトが入っていないか確認
            cellValue = ws.Cells(staffIndex + 2, dateIndex + i + 2).Value
            If cellValue = g_marks.dayShiftMark Or cellValue = g_marks.nightShiftMark Then
                ' 日勤または夜勤が既に割り当てられている場合、割り当て不可
                CheckFutureDays = False
                Exit Function
            End If
            ' 夜勤明けマークや休みマークの場合は問題なし
        End If
    Next i
End Function



Sub SortStaffByHoursWorked(ByRef staffList As Collection, ByRef hoursWorked() As Double, ByRef totalShiftsAssigned() As Integer)
    Dim sortedStaff As New Collection
    Dim i As Integer, j As Integer
    Dim minIndex As Integer

    ' スタッフリストをランダムにシャッフル
    ShuffleCollection staffList

    While staffList.count > 0
        minIndex = 1
        For i = 2 To staffList.count
            ' 勤務時間と総シフト数で比較
            If (hoursWorked(staffList(i)) < hoursWorked(staffList(minIndex))) Or _
            (hoursWorked(staffList(i)) = hoursWorked(staffList(minIndex)) And totalShiftsAssigned(staffList(i)) < totalShiftsAssigned(staffList(minIndex))) Then
                minIndex = i
            End If
        Next i
        sortedStaff.Add staffList(minIndex)
        staffList.Remove minIndex
    Wend

    ' ソートされたスタッフリストを元のコレクションにコピー
    For i = 1 To sortedStaff.count
        staffList.Add sortedStaff(i)
    Next i
End Sub

Sub ShuffleCollection(col As Collection)
    Dim i As Integer, j As Integer
    Dim temp As Variant
    Dim values() As Variant
    Dim count As Integer

    ' コレクションの要素を配列にコピー
    count = col.count
    If count = 0 Then Exit Sub ' 空のコレクションの場合は何もしない
    ReDim values(1 To count)
    For i = 1 To count
        values(i) = col(i)
    Next i

    ' 配列の要素をシャッフル
    For i = count To 2 Step -1
        Randomize ' ランダムシードを初期化
        j = Int(Rnd() * i) + 1
        temp = values(i)
        values(i) = values(j)
        values(j) = temp
    Next i

    ' 新しいコレクションを作成してシャッフルされた配列の要素を追加
    Dim newCol As Collection
    Set newCol = New Collection
    For i = 1 To count
        newCol.Add values(i)
    Next i

    ' 元のコレクションに新しいコレクションを置き換え
    For i = 1 To count
        col.Remove 1
    Next i
    For i = 1 To count
        col.Add newCol(i)
    Next i
End Sub

Sub SortStaffByRemainingNightShifts(ByRef staffList As Collection, ByRef nightShiftCountWorked() As Integer, ByRef staffNightShiftMax() As Integer)
    Dim sortedStaff As New Collection
    Dim i As Integer, j As Integer
    Dim maxIndex As Integer
    Dim remainingNightShifts As Integer

    ' スタッフリストをランダムにシャッフル
    ShuffleCollection staffList

    While staffList.count > 0
        maxIndex = 1
        For i = 2 To staffList.count
            If (staffNightShiftMax(staffList(i)) - nightShiftCountWorked(staffList(i))) > (staffNightShiftMax(staffList(maxIndex)) - nightShiftCountWorked(staffList(maxIndex))) Then
                maxIndex = i
            End If
        Next i
        sortedStaff.Add staffList(maxIndex)
        staffList.Remove maxIndex
    Wend

    ' ソートされたスタッフリストを元のコレクションにコピー
    For i = 1 To sortedStaff.count
        staffList.Add sortedStaff(i)
    Next i
End Sub

Function IsHoliday(currentDate As Date, holidays As Collection) As Boolean
    Dim holidayDate As Variant
    IsHoliday = False
    For Each holidayDate In holidays
        If holidayDate = currentDate Then
            IsHoliday = True
            Exit Function
        End If
    Next holidayDate
End Function

' 新しい関数を追加
Function IsAvailableForNightShift(staffIndex As Integer, currentDate As Date) As Boolean
    Dim dayOfWeek As Integer
    dayOfWeek = Weekday(currentDate, vbSunday)
        
    If staffFridaySaturdayOnlyNightShift(staffIndex) = "1" Then
        ' スタッフが金・土のみ夜勤可能な場合
        If dayOfWeek = vbFriday Or dayOfWeek = vbSaturday Then
            IsAvailableForNightShift = True
        Else
            IsAvailableForNightShift = False
        End If
    Else
        ' スタッフが制限なく夜勤可能な場合
        IsAvailableForNightShift = True
    End If
End Function


Function IsStaffHoliday(staffIndex As Integer, dateToCheck As Date) As Boolean
    Dim holidayDate As Variant
    IsStaffHoliday = False
    For Each holidayDate In staffHolidays(staffIndex)
        If holidayDate = dateToCheck Then
            IsStaffHoliday = True
            Exit Function
        End If
    Next holidayDate
End Function
