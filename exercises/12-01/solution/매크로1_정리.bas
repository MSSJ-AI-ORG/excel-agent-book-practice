Attribute VB_Name = "기록매크로"
Option Explicit

' 매크로1
' 매크로 기록: 2024-03-11
'
Sub 매크로1()
'
' 매크로1 매크로 - 머리글 굵게·색, 금액 표시 형식, 3행까지 틀 고정
'
    Range("A3:F3").Font.Bold = True
    With Range("A3:F3").Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Range("D4:D19").NumberFormat = "#,##0"
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 3
    ActiveWindow.FreezePanes = True
End Sub

