' ShiftSchedulerBeta.bas
' Version: Beta
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
    Dim staffMaxHours() As Double
    Dim staffDayShiftMax() As Integer
    Dim staffNightShiftMax() As Integer
    Dim staffCompatibility() As String
    Dim staffShifts() As Collection
    Dim staffHolidays() As Collection
    Dim i As Integer, j As Integer, k As Integer
    Dim currentDate As Date
    Dim dayOfWeek As Integer
    Dim dayShiftCount As Integer
    Dim nightShiftCount As Integer
    Dim dayShiftAssigned As Integer
    Dim nightShiftAssigned As Integer
    Dim hoursWorked() As Double
    Dim nightShiftCountWorked() As Integer
    Dim availableStaff As Collection
    Dim staffIndex As Variant
    Dim logFile As String
    Dim holidayDates() As Date
    Dim holidayDate As Variant

    logFile = ThisWorkbook.Path & "\ShiftScheduleLog.txt"
    Open logFile For Output As #1
    
    Print #1, "Step 1: Setting worksheet"
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' 入力シートからデータを取得
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
    Dim dayShiftMark As String
    Dim nightShiftMark As String
    Dim nightShiftAfterMark As String
    Dim holidayMark As String
    dayShiftMark = wsInput.Cells(1, 12).Value ' L1
    nightShiftMark = wsInput.Cells(2, 12).Value ' L2
    nightShiftAfterMark = wsInput.Cells(3, 12).Value ' L3
    holidayMark = wsInput.Cells(4, 12).Value ' L4
    
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
    ReDim staffCompatibility(1 To lastRow - 9)
    ReDim staffShifts(1 To lastRow - 9)
    ReDim staffHolidays(1 To lastRow - 9)
    ReDim hoursWorked(1 To lastRow - 9)
    ReDim nightShiftCountWorked(1 To lastRow - 9)

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
        ReDim Preserve staffCompatibility(1 To staffCount)
        ReDim Preserve staffShifts(1 To staffCount)
        ReDim Preserve staffHolidays(1 To staffCount)
        ReDim Preserve hoursWorked(1 To staffCount)
        ReDim Preserve nightShiftCountWorked(1 To staffCount)

        staffNames(staffCount) = wsInput.Cells(i, 1).Value
        If IsEmpty(wsInput.Cells(i, 2).Value) Then
            staffAttributes(staffCount) = ""
        Else
            staffAttributes(staffCount) = wsInput.Cells(i, 2).Value
        End If
        staffCompatibility(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 3).Value), wsInput.Cells(i, 3).Value, "")
        staffMaxHours(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 4).Value), wsInput.Cells(i, 4).Value, 0)
        staffDayShiftMax(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 5).Value), wsInput.Cells(i, 5).Value, 0)
        staffNightShiftMax(staffCount) = IIf(IsNumeric(wsInput.Cells(i, 6).Value), wsInput.Cells(i, 6).Value, 0)
        Set staffShifts(staffCount) = New Collection
        Set staffHolidays(staffCount) = New Collection
        
        Print #1, "Staff: " & staffNames(staffCount) & ", Attribute: " & staffAttributes(staffCount) & _
                    ", Compatibility: " & staffCompatibility(staffCount) & ", MaxHours: " & staffMaxHours(staffCount) & _
                    ", DayShiftMax: " & staffDayShiftMax(staffCount) & ", NightShiftMax: " & staffNightShiftMax(staffCount)
        
        For j = 7 To 16
            If IsDate(wsInput.Cells(i, j).Value) Then
                staffHolidays(staffCount).Add wsInput.Cells(i, j).Value
                Print #1, "  Holiday: " & wsInput.Cells(i, j).Value
            Else
                Print #1, "  Non-date value ignored: " & wsInput.Cells(i, j).Value
            End If
        Next j
NextStaff:
    Next i
    On Error GoTo ErrorHandler
    
    ' 公休日を取得
    Print #1, "Step 4: Getting public holidays"
    On Error GoTo HolidayDataError
    ReDim holidayDates(1 To 6)
    For i = 1 To 6
        If IsDate(wsInput.Cells(2, i + 3).Value) Then
            holidayDates(i) = wsInput.Cells(2, i + 3).Value
            Print #1, "  Public Holiday: " & holidayDates(i)
        Else
            holidayDates(i) = 0
        End If
    Next i

    ' 入力シートのD2:I2に公休日を追加
    For i = 1 To 6
        If IsDate(wsInput.Cells(2, i + 3).Value) Then
            holidayDates(i) = wsInput.Cells(2, i + 3).Value
            Print #1, "  Public Holiday: " & holidayDates(i)
        End If
    Next i

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

    Dim dates As Collection
    Set dates = New Collection

    currentDate = startDate
    Do While currentDate <= endDate
        dates.Add currentDate
        wsOutput.Cells(2, dates.count + 2).Value = Format(currentDate, "m/d(aaa)")
        currentDate = DateAdd("d", 1, currentDate)
    Loop
    On Error GoTo ErrorHandler

    ' スタッフ名の列を作成
    Print #1, "Step 8: Creating staff name column"
    On Error GoTo StaffNameColumnError
    For i = 1 To UBound(staffNames)
        If staffNames(i) <> "" Then
            wsOutput.Cells(i + 2, 1).Value = staffNames(i)
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
                For k = 1 To dates.count
                    If dates(k) = holidayDate Then
                        wsOutput.Cells(i + 2, k + 2).Value = holidayMark
                        wsOutput.Cells(i + 2, k + 2).Interior.Color = RGB(255, 255, 153) ' 薄い黄色
                        Exit For
                    End If
                Next k
            Next holidayDate
        End If
    Next i

    ' 公休日のセルの色を変更
    For k = 1 To dates.count
        currentDate = dates(k)
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

    ' 初期化
    ReDim hoursWorked(1 To staffCount)
    ReDim nightShiftCountWorked(1 To staffCount)

    ' 日毎のシフトカウント用変数を初期化
    Dim dayShiftDailyCount() As Integer
    Dim nightShiftDailyCount() As Integer
    Dim offShiftDailyCount() As Integer
    ReDim dayShiftDailyCount(1 To dates.count)
    ReDim nightShiftDailyCount(1 To dates.count)
    ReDim offShiftDailyCount(1 To dates.count)

    ' 1. 全ての日にリーダーを割り付ける
    Dim dayShiftLeaderAssigned As Boolean
    Dim nightShiftLeaderAssigned As Boolean
    
    For i = 1 To dates.count
        currentDate = dates(i)
        dayOfWeek = Weekday(currentDate, vbSunday)
        Print #1, "currentDate: " & currentDate & ", dayOfWeek: " & dayOfWeek
        
        If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
            dayShiftCount = weekdayDayShiftMin
            nightShiftCount = weekdayNightShiftMin
        Else
            dayShiftCount = weekendDayShiftMin
            nightShiftCount = weekendNightShiftMin
        End If

        For Each holidayDate In holidayDates
            If currentDate = holidayDate Then
                dayShiftCount = weekendDayShiftMin
                nightShiftCount = weekendNightShiftMin
                Exit For
            End If
        Next holidayDate

        Print #1, "dayShiftCount: " & dayShiftCount & ", nightShiftCount: " & nightShiftCount

        Set availableStaff = New Collection
        Dim uniqueStaff As Collection
        Set uniqueStaff = New Collection

        For j = 1 To UBound(staffNames)
            If staffNames(j) <> "" Then
                If wsOutput.Cells(j + 2, i + 2).Value = holidayMark Or wsOutput.Cells(j + 2, i + 2).Value = dayShiftMark Or wsOutput.Cells(j + 2, i + 2).Value = nightShiftMark Then
                    wsOutput.Cells(j + 2, i + 2).Value = wsOutput.Cells(j + 2, i + 2).Value ' Keep existing value
                ElseIf hoursWorked(j) < staffMaxHours(j) Then
                    On Error Resume Next
                    uniqueStaff.Add j, CStr(j)
                    If Err.Number = 0 Then
                        availableStaff.Add j
                        Print #1, "  availableStaff: " & staffNames(j)
                    End If
                    On Error GoTo 0
                End If
            End If
        Next j

        ' 勤務時間が少ないスタッフを優先するためにソート
        Call SortStaffByHoursWorked(availableStaff, hoursWorked)

        ' 利用可能なスタッフのリストをランダム化
        Call ShuffleCollection(availableStaff)
        
        dayShiftAssigned = 0
        nightShiftAssigned = 0
        Dim dayShiftLeaders As Integer
        Dim nightShiftLeaders As Integer
        dayShiftLeaders = 0
        nightShiftLeaders = 0
        
        ' 夜勤リーダーを割り当てる
        nightShiftLeaderAssigned = False
        For Each staffIndex In availableStaff
            If Not nightShiftLeaderAssigned And staffAttributes(staffIndex) = "2" And wsOutput.Cells(staffIndex + 2, i + 2).Value = "" Then
                If hoursWorked(staffIndex) + 16.5 <= staffMaxHours(staffIndex) And nightShiftCountWorked(staffIndex) < staffNightShiftMax(staffIndex) Then
                    ' ここで連続夜勤のチェックを緩和する
                    If CanAssignShift(wsOutput, staffIndex, i, "night", False, dayShiftMark, nightShiftMark) Then
                        If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "night", dayShiftMark, nightShiftMark) Then
                            Call AssignShift(wsOutput, staffIndex, i, nightShiftMark, 16.5, hoursWorked, staffShifts, staffNames, "night", staffAttributes, currentDate)
                            nightShiftCountWorked(staffIndex) = nightShiftCountWorked(staffIndex) + 1
                            nightShiftAssigned = nightShiftAssigned + 1
                            nightShiftDailyCount(i) = nightShiftDailyCount(i) + 1
                            nightShiftLeaderAssigned = True
                            nightShiftLeaders = nightShiftLeaders + 1
                            If i + 1 <= dates.count Then wsOutput.Cells(staffIndex + 2, i + 3).Value = nightShiftAfterMark
                            If i + 2 <= dates.count Then wsOutput.Cells(staffIndex + 2, i + 4).Value = holidayMark
                            Exit For ' リーダーが割り当てられたらループを終了
                        End If
                    End If
                End If
            End If
        Next staffIndex
        
        ' 日勤リーダーを割り当てる
        dayShiftLeaderAssigned = False
        For Each staffIndex In availableStaff
            If Not dayShiftLeaderAssigned And staffAttributes(staffIndex) = "2" And wsOutput.Cells(staffIndex + 2, i + 2).Value = "" Then
                If hoursWorked(staffIndex) + 8.5 <= staffMaxHours(staffIndex) Then
                    If CanAssignShift(wsOutput, staffIndex, i, "day", True, dayShiftMark, nightShiftMark) Then
                        If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "day", dayShiftMark, nightShiftMark) Then
                            Call AssignShift(wsOutput, staffIndex, i, dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, "day", staffAttributes, currentDate)
                            dayShiftAssigned = dayShiftAssigned + 1
                            dayShiftDailyCount(i) = dayShiftDailyCount(i) + 1
                            dayShiftLeaderAssigned = True
                            dayShiftLeaders = dayShiftLeaders + 1
                            Exit For ' リーダーが割り当てられたらループを終了
                        End If
                    End If
                End If
            End If
        Next staffIndex

        ' 日勤と夜勤のリーダー数を記録
        wsOutput.Cells(staffCount + 6, i + 2).Value = dayShiftLeaders
        wsOutput.Cells(staffCount + 7, i + 2).Value = nightShiftLeaders

        ' 夜勤リーダーが割り振られた後の人数を再計算
        nightShiftAssigned = nightShiftLeaders
    Next i
    
    ' 2. 各日の最低限の勤務を割り付ける
    For i = 1 To dates.count
        currentDate = dates(i)
        dayOfWeek = Weekday(currentDate, vbSunday)
        If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
            dayShiftCount = weekdayDayShiftMin
            nightShiftCount = weekdayNightShiftMin
        Else
            dayShiftCount = weekendDayShiftMin
            nightShiftCount = weekendNightShiftMin
        End If

        For Each holidayDate In holidayDates
            If currentDate = holidayDate Then
                dayShiftCount = weekendDayShiftMin
                nightShiftCount = weekendNightShiftMin
                Exit For
            End If
        Next holidayDate

        Set availableStaff = New Collection
        Dim availableLeaders As Collection
        Set availableLeaders = New Collection
        
        ' リーダー属性（1または2）を持つスタッフを優先的に追加
        For j = 1 To UBound(staffNames)
            If staffNames(j) <> "" Then
                If wsOutput.Cells(j + 2, i + 2).Value = "" And hoursWorked(j) < staffMaxHours(j) Then
                    If staffAttributes(j) = "1" Or staffAttributes(j) = "2" Then
                        availableLeaders.Add j
                    Else
                        availableStaff.Add j
                    End If
                End If
            End If
        Next j

        ' 勤務時間が少ないスタッフを優先するためにソート
        Call SortStaffByHoursWorked(availableLeaders, hoursWorked)
        Call SortStaffByHoursWorked(availableStaff, hoursWorked)

        ' 利用可能なスタッフのリストをランダム化
        Call ShuffleCollection(availableLeaders)
        Call ShuffleCollection(availableStaff)
        
        dayShiftAssigned = dayShiftDailyCount(i)
        nightShiftAssigned = nightShiftDailyCount(i) ' 既にリーダーシフトが割り当てられているため初期値を設定
        
        ' 最低限のシフトをリーダー属性を持つスタッフに割り当てる
        For Each staffIndex In availableLeaders
            If wsOutput.Cells(staffIndex + 2, i + 2).Value = "" Then
                If dayShiftAssigned < dayShiftCount And staffShifts(staffIndex).count < staffDayShiftMax(staffIndex) Then
                    If hoursWorked(staffIndex) + 8.5 <= staffMaxHours(staffIndex) Then
                        If CanAssignShift(wsOutput, staffIndex, i, "day", True, dayShiftMark, nightShiftMark) Then
                            If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "day", dayShiftMark, nightShiftMark) Then
                                Call AssignShift(wsOutput, staffIndex, i, dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, "day", staffAttributes, currentDate)
                                dayShiftAssigned = dayShiftAssigned + 1
                                dayShiftDailyCount(i) = dayShiftDailyCount(i) + 1
                            Else
                                Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " has incompatible staff in day shift"
                            End If
                        Else
                            Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned day shift due to max continuous day shift limit"
                        End If
                    Else
                        Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned day shift due to max hours limit"
                    End If
                ElseIf nightShiftAssigned < nightShiftCount And nightShiftCountWorked(staffIndex) < staffNightShiftMax(staffIndex) Then
                    If hoursWorked(staffIndex) + 16.5 <= staffMaxHours(staffIndex) Then
                        ' ここで連続夜勤のチェックを緩和する
                        If CanAssignShift(wsOutput, staffIndex, i, "night", False, dayShiftMark, nightShiftMark) Then
                            If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "night", dayShiftMark, nightShiftMark) Then
                                Call AssignShift(wsOutput, staffIndex, i, nightShiftMark, 16.5, hoursWorked, staffShifts, staffNames, "night", staffAttributes, currentDate)
                                nightShiftCountWorked(staffIndex) = nightShiftCountWorked(staffIndex) + 1
                                nightShiftAssigned = nightShiftAssigned + 1
                                nightShiftDailyCount(i) = nightShiftDailyCount(i) + 1
                                If i + 1 <= dates.count Then
                                    If wsOutput.Cells(staffIndex + 2, i + 3).Value <> holidayMark Then wsOutput.Cells(staffIndex + 2, i + 3).Value = nightShiftAfterMark
                                End If
                                If i + 2 <= dates.count Then
                                    If wsOutput.Cells(staffIndex + 2, i + 4).Value <> holidayMark Then wsOutput.Cells(staffIndex + 2, i + 4).Value = holidayMark
                                End If
                            Else
                                Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " has incompatible staff in night shift"
                            End If
                        Else
                            Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned night shift due to max continuous night shift limit"
                        End If
                    Else
                        Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned night shift due to max hours limit"
                    End If
                End If
            End If
            If dayShiftAssigned >= dayShiftCount And nightShiftAssigned >= nightShiftCount Then Exit For
        Next staffIndex

        ' 最低限のシフトが足りない場合、他のスタッフも含めて割り当てる
        For Each staffIndex In availableStaff
            If wsOutput.Cells(staffIndex + 2, i + 2).Value = "" Then
                If dayShiftAssigned < dayShiftCount And staffShifts(staffIndex).count < staffDayShiftMax(staffIndex) Then
                    If hoursWorked(staffIndex) + 8.5 <= staffMaxHours(staffIndex) Then
                        If CanAssignShift(wsOutput, staffIndex, i, "day", True, dayShiftMark, nightShiftMark) Then
                            If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "day", dayShiftMark, nightShiftMark) Then
                                Call AssignShift(wsOutput, staffIndex, i, dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, "day", staffAttributes, currentDate)
                                dayShiftAssigned = dayShiftAssigned + 1
                                dayShiftDailyCount(i) = dayShiftDailyCount(i) + 1
                            Else
                                Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " has incompatible staff in day shift"
                            End If
                        Else
                            Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned day shift due to max continuous day shift limit"
                        End If
                    Else
                        Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned day shift due to max hours limit"
                    End If
                ElseIf nightShiftAssigned < nightShiftCount And nightShiftCountWorked(staffIndex) < staffNightShiftMax(staffIndex) Then
                    If hoursWorked(staffIndex) + 16.5 <= staffMaxHours(staffIndex) Then
                        ' ここで連続夜勤のチェックを緩和する
                        If CanAssignShift(wsOutput, staffIndex, i, "night", False, dayShiftMark, nightShiftMark) Then
                            If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "night", dayShiftMark, nightShiftMark) Then
                                Call AssignShift(wsOutput, staffIndex, i, nightShiftMark, 16.5, hoursWorked, staffShifts, staffNames, "night", staffAttributes, currentDate)
                                nightShiftCountWorked(staffIndex) = nightShiftCountWorked(staffIndex) + 1
                                nightShiftAssigned = nightShiftAssigned + 1
                                nightShiftDailyCount(i) = nightShiftDailyCount(i) + 1
                                If i + 1 <= dates.count Then
                                    If wsOutput.Cells(staffIndex + 2, i + 3).Value <> holidayMark Then wsOutput.Cells(staffIndex + 2, i + 3).Value = nightShiftAfterMark
                                End If
                                If i + 2 <= dates.count Then
                                    If wsOutput.Cells(staffIndex + 2, i + 4).Value <> holidayMark Then wsOutput.Cells(staffIndex + 2, i + 4).Value = holidayMark
                                End If
                            Else
                                Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " has incompatible staff in night shift"
                            End If
                        Else
                            Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned night shift due to max continuous night shift limit"
                        End If
                    Else
                        Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned night shift due to max hours limit"
                    End If
                End If
            End If
            If dayShiftAssigned >= dayShiftCount And nightShiftAssigned >= nightShiftCount Then Exit For
        Next staffIndex

        ' デバッグメッセージを追加
        Print #1, currentDate & ": Additional Day Shift Assigned: " & dayShiftAssigned - dayShiftDailyCount(i) & " Additional Night Shift Assigned: " & nightShiftAssigned - nightShiftDailyCount(i)
    Next i

    ' 3. 各日の余剰勤務を割り付ける
    For i = 1 To dates.count
        currentDate = dates(i)
        dayOfWeek = Weekday(currentDate, vbSunday)
        If dayOfWeek >= vbMonday And dayOfWeek <= vbFriday Then
            dayShiftCount = weekdayDayShiftMax
            nightShiftCount = weekdayNightShiftMax
        Else
            dayShiftCount = weekendDayShiftMax
            nightShiftCount = weekendNightShiftMax
        End If

        For Each holidayDate In holidayDates
            If currentDate = holidayDate Then
                dayShiftCount = weekendDayShiftMax
                nightShiftCount = weekendNightShiftMax
                Exit For
            End If
        Next holidayDate

        Set availableStaff = New Collection
        For j = 1 To UBound(staffNames)
            If staffNames(j) <> "" Then
                If wsOutput.Cells(j + 2, i + 2).Value = "" And hoursWorked(j) < staffMaxHours(j) Then
                    availableStaff.Add j
                End If
            End If
        Next j

        ' 勤務時間が少ないスタッフを優先するためにソート
        Call SortStaffByHoursWorked(availableStaff, hoursWorked)

        ' 利用可能なスタッフのリストをランダム化
        Call ShuffleCollection(availableStaff)
        
        dayShiftAssigned = dayShiftDailyCount(i)
        nightShiftAssigned = nightShiftDailyCount(i)
        
   ' 余剰シフトの割り振り
For Each staffIndex In availableStaff
    If wsOutput.Cells(staffIndex + 2, i + 2).Value = "" Then
        If dayShiftAssigned < dayShiftCount And staffShifts(staffIndex).count < staffDayShiftMax(staffIndex) Then
            If hoursWorked(staffIndex) + 8.5 <= staffMaxHours(staffIndex) Then
                If CanAssignShift(wsOutput, staffIndex, i, "day", True, dayShiftMark, nightShiftMark) Then
                    If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "day", dayShiftMark, nightShiftMark) Then
                        Call AssignShift(wsOutput, staffIndex, i, dayShiftMark, 8.5, hoursWorked, staffShifts, staffNames, "day", staffAttributes, currentDate)
                        dayShiftAssigned = dayShiftAssigned + 1
                        dayShiftDailyCount(i) = dayShiftDailyCount(i) + 1
                    Else
                        Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " has incompatible staff in day shift"
                    End If
                Else
                    Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned day shift due to max continuous day shift limit"
                End If
            Else
                Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned day shift due to max hours limit"
            End If
        ElseIf nightShiftAssigned < nightShiftCount And nightShiftCountWorked(staffIndex) < staffNightShiftMax(staffIndex) Then
            If hoursWorked(staffIndex) + 16.5 <= staffMaxHours(staffIndex) Then
                ' ここで連続夜勤のチェックを緩和する
                If CanAssignShift(wsOutput, staffIndex, i, "night", False, dayShiftMark, nightShiftMark) Then
                    If Not HasIncompatibleStaff(wsOutput, staffCompatibility, staffIndex, i, "night", dayShiftMark, nightShiftMark) Then
                        Call AssignShift(wsOutput, staffIndex, i, nightShiftMark, 16.5, hoursWorked, staffShifts, staffNames, "night", staffAttributes, currentDate)
                        nightShiftCountWorked(staffIndex) = nightShiftCountWorked(staffIndex) + 1
                        nightShiftAssigned = nightShiftAssigned + 1
                        nightShiftDailyCount(i) = nightShiftDailyCount(i) + 1
                        If i + 1 <= dates.count Then
                            If wsOutput.Cells(staffIndex + 2, i + 3).Value <> holidayMark Then wsOutput.Cells(staffIndex + 2, i + 3).Value = nightShiftAfterMark
                        End If
                        If i + 2 <= dates.count Then
                            If wsOutput.Cells(staffIndex + 2, i + 4).Value <> holidayMark Then wsOutput.Cells(staffIndex + 2, i + 4).Value = holidayMark
                        End If
                    Else
                        Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " has incompatible staff in night shift"
                    End If
                Else
                    Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned night shift due to max continuous night shift limit"
                End If
            Else
                Print #1, currentDate & ":  Staff " & staffNames(staffIndex) & " cannot be assigned night shift due to max hours limit"
            End If
        End If
    End If
    If dayShiftAssigned >= dayShiftCount And nightShiftAssigned >= nightShiftCount Then Exit For
Next staffIndex

        ' デバッグメッセージを追加
        Print #1, currentDate & ": Additional Day Shift Assigned: " & dayShiftAssigned - dayShiftDailyCount(i) & " Additional Night Shift Assigned: " & nightShiftAssigned - nightShiftDailyCount(i)
    Next i

    On Error GoTo ErrorHandler

    ' 各スタッフの勤務時間を最終日に記載
    Print #1, "Step 11: Recording hours worked"
    On Error GoTo HoursRecordingError
    For i = 1 To UBound(staffNames)
        If staffNames(i) <> "" Then
            wsOutput.Cells(i + 2, dates.count + 3).Value = hoursWorked(i)
        End If
    Next i
    On Error GoTo ErrorHandler

    ' 日毎のシフトカウントを追加
    Print #1, "Step 12: Recording daily shift counts"
    On Error GoTo ShiftCountError
    wsOutput.Cells(staffCount + 3, 1).Value = "日勤"
    wsOutput.Cells(staffCount + 4, 1).Value = "夜勤"
    wsOutput.Cells(staffCount + 5, 1).Value = "休み"
    wsOutput.Cells(staffCount + 6, 1).Value = "日勤リーダー"
    wsOutput.Cells(staffCount + 7, 1).Value = "夜勤リーダー"
    For i = 1 To dates.count
        wsOutput.Cells(staffCount + 3, i + 2).Value = dayShiftDailyCount(i)
        wsOutput.Cells(staffCount + 4, i + 2).Value = nightShiftDailyCount(i)
        wsOutput.Cells(staffCount + 5, i + 2).Value = staffCount - dayShiftDailyCount(i) - nightShiftDailyCount(i)
    Next i
    On Error GoTo ErrorHandler

    ' 空欄セルを休み("X")で埋める
    Print #1, "Step 13: Filling empty cells with 'X'"
    For i = 1 To UBound(staffNames)
        For j = 1 To dates.count
            If wsOutput.Cells(i + 2, j + 2).Value = "" Then
                wsOutput.Cells(i + 2, j + 2).Value = holidayMark
            End If
        Next j
    Next i
    On Error GoTo ErrorHandler

    Print #1, "Shift schedule created successfully"
    Close #1
    MsgBox "Shift schedule created successfully"

    Exit Sub

DataError:
    Print #1, "Data Error: " & Err.Description & " at line " & Erl & " (Step 2)"
    Resume Next

StaffDataError:
    Print #1, "Staff Data Error: " & Err.Description & " at line " & Erl & " (Step 3)"
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

HolidayDataError:
    Print #1, "Holiday Data Error: " & Err.Description & " at line " & Erl & " (Step 4)"
    Resume Next

ShiftScheduleError:
    Print #1, "Shift Schedule Error: " & Err.Description & " at line " & Erl & " (Step 10)"
    Resume Next

HoursRecordingError:
    Print #1, "Hours Recording Error: " & Err.Description & " at line " & Erl & " (Step 11)"
    Resume Next

ShiftCountError:
    Print #1, "Shift Count Error: " & Err.Description & " at line " & Erl & " (Step 12)"
    Resume Next

ErrorHandler:
    Print #1, "Error: " & Err.Description & " at line " & Erl
    Close #1
End Sub

Sub ShuffleCollection(col As Collection)
    Dim i As Integer, j As Integer
    Dim temp As Variant
    Dim values() As Variant
    Dim count As Integer

    ' コレクションの要素を配列にコピー
    count = col.count
    ReDim values(1 To count)
    For i = 1 To count
        values(i) = col(i)
    Next i

    ' 配列の要素をシャッフル
    For i = count To 2 Step -1
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

Sub AssignShift(ws As Worksheet, ByVal staffIndex As Integer, ByVal dateIndex As Integer, ByVal shiftMark As String, ByVal hours As Double, _
                ByRef hoursWorked() As Double, ByRef staffShifts() As Collection, ByRef staffNames() As String, ByVal shiftType As String, _
                ByRef staffAttributes() As String, ByVal currentDate As Date)
    ws.Cells(staffIndex + 2, dateIndex + 2).Value = shiftMark
    hoursWorked(staffIndex) = hoursWorked(staffIndex) + hours
    staffShifts(staffIndex).Add shiftType
    Print #1, currentDate & ":  Assigned " & shiftType & " shift to: " & staffNames(staffIndex) & ", Total Hours Worked: " & hoursWorked(staffIndex)

    ' リーダーのカウントを更新
    If shiftType = "day" And staffAttributes(staffIndex) = "2" Then
        ws.Cells(UBound(staffNames) + 6, dateIndex + 2).Value = ws.Cells(UBound(staffNames) + 6, dateIndex + 2).Value + 1
    ElseIf shiftType = "night" And staffAttributes(staffIndex) = "2" Then
        ws.Cells(UBound(staffNames) + 7, dateIndex + 2).Value = ws.Cells(UBound(staffNames) + 7, dateIndex + 2).Value + 1
    End If
End Sub

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

Function CanAssignShift(ws As Worksheet, ByVal staffIndex As Integer, ByVal dateIndex As Integer, ByVal shiftType As String, ByVal strictCheck As Boolean, ByVal dayShiftMark As String, ByVal nightShiftMark As String) As Boolean
    Dim i As Integer
    Dim maxContinuousShifts As Integer
    maxContinuousShifts = 5 ' 最大連続日数をデフォルトの5日に設定
    Dim continuousShifts As Integer
    continuousShifts = 0

    CanAssignShift = True

    ' 連続日勤のチェック
    If shiftType = "day" Then
        ' 指定日の前後を含む連続日勤チェック
        For i = dateIndex To dateIndex - 5 Step -1
            If i < 1 Then Exit For
            If ws.Cells(staffIndex + 2, i + 2).Value = dayShiftMark Then
                continuousShifts = continuousShifts + 1
            ElseIf ws.Cells(staffIndex + 2, i + 2).Value <> "" Then ' 空白でないセルはすべてリセットとみなす
                Exit For
            End If
        Next i
        
        For i = dateIndex + 1 To dateIndex + 5
            If i > ws.Cells(2, Columns.count).End(xlToLeft).Column - 2 Then Exit For
            If ws.Cells(staffIndex + 2, i + 2).Value = dayShiftMark Then
                continuousShifts = continuousShifts + 1
            ElseIf ws.Cells(staffIndex + 2, i + 2).Value <> "" Then ' 空白でないセルはすべてリセットとみなす
                Exit For
            End If
        Next i

        If continuousShifts >= maxContinuousShifts Then
            CanAssignShift = False
            Exit Function
        End If
    End If

    ' 連続夜勤のチェック
    If shiftType = "night" Then
        continuousShifts = 0
        ' 指定日の前後を含む連続夜勤チェック
        For i = dateIndex To dateIndex - 5 Step -1
            If i < 1 Then Exit For
            If ws.Cells(staffIndex + 2, i + 2).Value = nightShiftMark Then
                continuousShifts = continuousShifts + 1
            ElseIf ws.Cells(staffIndex + 2, i + 2).Value <> "" Then
                Exit For
            End If
        Next i
        
        For i = dateIndex + 1 To dateIndex + 5
            If i > ws.Cells(2, Columns.count).End(xlToLeft).Column - 2 Then Exit For
            If ws.Cells(staffIndex + 2, i + 2).Value = nightShiftMark Then
                continuousShifts = continuousShifts + 1
            ElseIf ws.Cells(staffIndex + 2, i + 2).Value <> "" Then
                Exit For
            End If
        Next i

        If continuousShifts >= maxContinuousShifts Then
            CanAssignShift = False
            Exit Function
        End If
    End If
    
    ' 夜勤明けの次の日に夜勤を割り当てる場合のチェック
    If shiftType = "night" And strictCheck Then
        If dateIndex > 1 And ws.Cells(staffIndex + 2, dateIndex + 1).Value = nightShiftAfterMark Then
            If ws.Cells(staffIndex + 2, dateIndex).Value = nightShiftMark Then
                CanAssignShift = False
                Exit Function
            End If
        End If
    End If
End Function



Sub SortStaffByHoursWorked(ByRef staffList As Collection, ByRef hoursWorked() As Double)
    Dim sortedStaff As New Collection
    Dim i As Integer, j As Integer
    Dim temp As Variant
    Dim minIndex As Integer

    While staffList.count > 0
        minIndex = 1
        For i = 2 To staffList.count
            If hoursWorked(staffList(i)) < hoursWorked(staffList(minIndex)) Then
                minIndex = i
            End If
        Next i
        sortedStaff.Add staffList(minIndex)
        staffList.Remove minIndex
    Wend

    ' Copy sorted elements back to original collection
    For i = 1 To sortedStaff.count
        staffList.Add sortedStaff(i)
    Next i
End Sub

