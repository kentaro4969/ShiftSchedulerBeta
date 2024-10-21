ShiftSchedulerBeta
A VBA script for generating and managing shift schedules. This is an updated version with additional features for better control over scheduling parameters. Feedback and contributions are welcome.

Features
Generate shift schedules based on input parameters
Manage and adjust shifts for employees, including day shifts, night shifts, and leadership roles
Customize maximum consecutive day shifts per staff member
Handle specific night shift scheduling, including Friday/Saturday-only night shift staff
Prevents assigning incompatible staff members to the same shift
User-friendly interface for easy modifications
Version History
Beta 1.2 (New Update)
New Customizable Maximum Consecutive Day Shifts: Added support for setting individual maximum consecutive day shifts per staff member (F列).
Friday/Saturday-Only Night Shifts: Staff who can only work Friday and Saturday night shifts are now supported (E列).
Enhanced Holiday Handling: Fixed issues related to public holidays and staff-specific holidays, ensuring no conflicts with night shifts or incompatible shifts.
Improved Night Shift Allocation: Adjustments to ensure consecutive night shifts are correctly assigned and calculated based on the staff's suitability and working hours.
Previous Month's Data Consideration: Integrated logic to handle the staff's previous month shift data, affecting the current month's schedule.
Beta 1.1
Added support for attributes 3 and 4.
Attribute 4 (head nurse) is assigned to weekday day shifts for ward management.
Attribute 3 (deputy head nurse) manages the ward on weekdays when the head nurse is absent.
Beta 1.0
Initial release with basic shift scheduling features.
Getting Started
To use this script, follow these steps:

1. Download the script
Download the ShiftSchedulerBeta.bas file from the repository.

2. Import the script into Excel
Open Excel and press Alt + F11 to open the VBA editor. In the VBA editor, go to File > Import File... and select the ShiftSchedulerBeta.bas file.

3. Run the script
Close the VBA editor. In Excel, press Alt + F8, select the ShiftSchedulerBeta macro, and click Run.

Usage
Setup: Define Employees and Their Availability
Download the sample InputSheet: ShiftSchedulerBeta.xlsm.
Fill in the details as described below.
Set the parameters for the shifts.
Inputシートの使い方
基本情報
病棟名: セル B1
開始日: セル B2 (yyyy/mm/dd形式)
平日日勤最大人数: セル B3
平日日勤最小人数: セル C3
平日夜勤最大人数: セル B4
平日夜勤最小人数: セル C4
週末日勤最大人数: セル B5
週末日勤最小人数: セル C5
週末夜勤最大人数: セル B6
週末夜勤最小人数: セル C6
日勤リーダー最小人数: セル B7
夜勤リーダー最小人数: セル B8
シフトマーク
昼勤マーク: セル P1
夜勤マーク: セル P2
夜勤後のマーク: セル P3
休暇マーク: セル P4
公休日 (E2:J2)
セル E2 からM2 までに公休日を入力 (yyyy/mm/dd形式)
スタッフ情報
名前: セル A10 以降
属性: セル B10 以降
師長(平日日勤、休日は原則休み)は"4"
主任(師長がいない日は日勤)は"3"
リーダーの場合は "2"
一般スタッフの場合は "1"
独り立ちしていないスタッフは空欄
2連続夜勤希望者(入明入明希望):C10以降 
相性: セル D10 以降 (組ませたくないスタッフ同士がいる場合は1-9の同じ数字を入力)
金・土夜勤専用スタッフ: セル E列（値が "1" の場合、金・土のみ夜勤可能）
最大連続日勤日数: セル F列 (最大連続日勤日数の指定が可能)
最大勤務時間: セル H10 以降 (時間数)
最大日勤回数: セル I10 以降
最大夜勤回数: セル J10 以降
希望休: セル K10 から T10 まで (yyyy/mm/dd形式で最大10日まで)
前月末夜勤情報: セル U列 (夜勤:1, 夜勤明け:2)
前月末連続日勤情報: 前月末の連続日勤数(日数)
Generate Schedule
Run the script to generate the initial shift schedule. Review and adjust the schedule as necessary.

Contribution
Since this is a beta version, feedback and contributions are highly appreciated. If you encounter any issues or have suggestions for improvement, please open an issue or submit a pull request.

License
This script is provided under a non-commercial use license. See the script comments for more details.
