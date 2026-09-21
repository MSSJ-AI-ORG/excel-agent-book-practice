Attribute VB_Name = "새매크로"
Option Explicit

' 대장의 금액을 부서별로 더해 새 파일로 저장한다. 부서 이름은 공백을 모두 빼고 댄다(「영업 1팀」=「영업1팀」 - 용어사전).
' 읽는 곳: 대장 B열(부서)·D열(금액), 4행부터 A열 마지막 데이터 행까지. 쓰는 곳: 새 통합 문서 하나. 이 파일은 바꾸지 않는다.
Sub 부서별합계저장()
    Const FIRST_ROW As Long = 4
    Const XLSX_FORMAT As Long = 51
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dept As String
    Dim sums As Object
    Dim key As Variant
    Dim outWb As Workbook
    Dim outRow As Long
    Dim savePath As String

    On Error GoTo 멈춤

    Set ws = ThisWorkbook.Worksheets("대장")
    Set sums = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = FIRST_ROW To lastRow
        dept = Replace(CStr(ws.Cells(r, 2).Value), " ", "")
        If dept <> "" Then
            sums(dept) = sums(dept) + ws.Cells(r, 4).Value
        End If
    Next r

    savePath = 저장경로()

    Set outWb = Workbooks.Add
    outWb.Worksheets(1).Range("A1:B1").Value = Array("부서", "합계")
    outRow = 2
    For Each key In sums.Keys
        outWb.Worksheets(1).Cells(outRow, 1).Value = key
        outWb.Worksheets(1).Cells(outRow, 2).Value = sums(key)
        outRow = outRow + 1
    Next key
    outWb.SaveAs Filename:=savePath, FileFormat:=XLSX_FORMAT
    outWb.Close SaveChanges:=False
    Exit Sub

멈춤:
    MsgBox "오류로 멈췄습니다: " & Err.Number & " " & Err.Description, vbCritical, "부서별합계저장"
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
End Sub

' 저장할 파일 경로 - 이 파일과 같은 폴더의 「부서별합계_YYYYMMDD_HHMM.xlsx」.
Function 저장경로() As String
    저장경로 = ThisWorkbook.path & "\부서별합계_" & Format(Now, "yyyymmdd_hhnn") & ".xlsx"
End Function

