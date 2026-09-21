Attribute VB_Name = "위험매크로"
Option Explicit

' 월말에 한 번 돌린다.
' 하는 일: 「백업_0402」 시트를 지우고, 이 파일을 「마감_YYYYMM.xlsx」로 새로 저장하고, 같은 폴더의 「임시.xlsx」를 지운다.
' 셋 다 되돌릴 수 없다. 그래서 하나씩 무엇이 사라지는지 보여 주고 묻는다. 경고(DisplayAlerts)는 끄지 않는다.
Sub 월말마감()
    Dim path As String
    Dim tmpPath As String
    Dim answer As VbMsgBoxResult

    On Error GoTo 멈춤

    path = ThisWorkbook.path & "\마감_" & Format(Date, "yyyymm") & ".xlsx"
    tmpPath = ThisWorkbook.path & "\임시.xlsx"

    answer = MsgBox("「백업_0402」 시트를 지웁니다. 지운 시트는 되돌릴 수 없습니다." & vbCrLf & "진행할까요?", vbYesNo + vbExclamation, "월말마감")
    If answer <> vbYes Then Exit Sub
    Sheets("백업_0402").Delete

    If Dir(path) <> "" Then
        answer = MsgBox(path & " 파일이 이미 있습니다. 덮어쓸까요?", vbYesNo + vbExclamation, "월말마감")
        If answer <> vbYes Then Exit Sub
    End If
    ThisWorkbook.SaveAs Filename:=path, FileFormat:=51

    If Dir(tmpPath) <> "" Then
        answer = MsgBox(tmpPath & " 파일을 지웁니다. 휴지통에 가지 않습니다." & vbCrLf & "진행할까요?", vbYesNo + vbExclamation, "월말마감")
        If answer = vbYes Then Kill tmpPath
    End If
    Exit Sub

멈춤:
    ' 경고를 끄지 않았으므로 되돌릴 설정은 없다. 무엇 때문에 멈췄는지만 알린다.
    MsgBox "오류로 멈췄습니다: " & Err.Number & " " & Err.Description, vbCritical, "월말마감"
End Sub

