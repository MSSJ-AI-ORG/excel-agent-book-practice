Attribute VB_Name = "깨진매크로"
Option Explicit

' 매달 돌리던 것. 언제부터인지 멈춘다.
' ── 고친 기록 ──
' 언제: 2026-09-21
' 무엇이 깨져 있었나: 없는 시트 「매출대장」을 잡다가 오류 9 로 멈췄다. 지우려던 「백업_0401」 시트도 이 파일에 없다.
' 무엇을 고쳤나: 「매출대장」 → 「대장」. 백업_0401 을 지우는 줄은 주석으로 막았다. 4~19행 고정 → A열 마지막 데이터 행까지.
' 확인은 어떻게 했나: 사본에서 돌려 D열이 3,000,000 을 넘는 네 행에만 「확인」이 적혔고, 시트 일곱 개가 그대로인 것을 봤다.
Sub 월말정리()
    Dim i As Long
    Dim lastRow As Long
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("대장")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        If ws.Cells(i, 4).Value > 3000000 Then
            ws.Cells(i, 6).Value = "확인"
        End If
    Next i

    ' 백업_0401 시트는 이 파일에 없다(있는 것은 백업_0402). 같은 백업인지 모르고, 지우면 되돌릴 수 없다.
    ' 사람이 정할 때까지 지우는 줄을 막아 둔다.
    'Application.DisplayAlerts = False
    'ThisWorkbook.Sheets("백업_0401").Delete
    'Application.DisplayAlerts = True
End Sub

Sub 요약만들기()
    Dim r As Long
    Dim total As Currency

    For r = 4 To 19
        total = total + Sheets("대장").Cells(r, 4)
    Next r

    Sheets("정산").Range("F1").Value = total
End Sub

