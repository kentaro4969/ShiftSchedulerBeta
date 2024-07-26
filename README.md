ShiftSchedulerBeta
A VBA script for generating and managing shift schedules. This is a beta version and still under development. Feedback and contributions are welcome.

Features
Generate shift schedules based on input parameters
Manage and adjust shifts for employees
User-friendly interface for easy modifications
Version History
Beta 1.0
Initial release with basic shift scheduling features
Getting Started
To use this script, follow these steps:

Download the script:

Download the ShiftSchedulerBeta.bas file from the repository.
Import the script into Excel:

Open Excel and press Alt + F11 to open the VBA editor.
In the VBA editor, go to File > Import File... and select the ShiftSchedulerBeta.bas file.
Run the script:

Close the VBA editor.
In Excel, press Alt + F8, select the ShiftSchedulerBeta macro, and click Run.
Usage
Setup:
Define the employees and their availability:

Download the sample InputSheet: ShiftSchedulerBeta.xlsm
Fill in the details as described below.
Set the parameters for the shifts:

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
昼勤マーク: セル L1
夜勤マーク: セル L2
夜勤後のマーク: セル L3
休暇マーク: セル L4

公休日 (任意)
セル D2 から I2 までに公休日を入力 (yyyy/mm/dd形式)

スタッフ情報
名前: セル A10 以降
属性: セル B10 以降 (リーダーの場合は "2"、一般スタッフの場合は "1" または空白)
相性: セル C10 以降 (相性が悪いスタッフ同士がいる場合は1-9の同じ数字を入力)
最大勤務時間: セル D10 以降 (時間数)
最大日勤回数: セル E10 以降
最大夜勤回数: セル F10 以降
希望休: セル G10 から P10 まで (yyyy/mm/dd形式で最大10日まで)

Generate Schedule:
Run the script to generate the initial shift schedule.
Review and adjust the schedule as necessary.
Contribution
Since this is a beta version, feedback and contributions are highly appreciated. If you encounter any issues or have suggestions for improvement, please open an issue or submit a pull request.

License
This script is provided under a non-commercial use license. See the script comments for more details.
