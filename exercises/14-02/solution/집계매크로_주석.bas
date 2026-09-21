Attribute VB_Name = "집계매크로"
Option Explicit

' 부서별 합계를 내고 목표와 견준다.
' 대상: 대장 B4:D19 / 기준 A3:B5 / 결과는 정산 G열
' 읽는 곳: 정산 A3:A5(부서 이름) · 대장 B4:B19(부서)와 D4:D19(금액) · 기준 B3:B5(목표). 쓰는 곳: 정산 G3:G5 뿐.
Sub 부서별집계()
    Dim i As Long, j As Long
    Dim dept As String
    Dim amount As Currency
    Dim target As Currency

    ' 정산 3~5행의 부서를 하나씩 본다 - 3번. 부서가 넷이 되면 넷째는 안 센다.
    For j = 3 To 5
        dept = Sheets("정산").Cells(j, 1).Value
        amount = 0

        ' 대장 4~19행을 훑는다 - 부서마다 16번, 모두 48번. 20행부터 들어온 거래는 안 센다.
        For i = 4 To 19
            ' 이 거래가 지금 부서의 것인가. Trim 은 앞뒤 공백만 뺀다 - 「영업 1팀」처럼 가운데 공백이 든 거래는 빠진다.
            If Trim(Sheets("대장").Cells(i, 2).Value) = dept Then
                amount = amount + Sheets("대장").Cells(i, 4).Value
            End If
        Next i

        target = Sheets("기준").Cells(j, 2).Value

        ' 목표를 채웠나(같으면 달성). 목표는 기준 시트의 같은 행 번호에서 읽는다 - 두 시트의 부서 순서가 같아야 맞다.
        If amount >= target Then
            Sheets("정산").Cells(j, 7).Value = "달성"
        Else
            Sheets("정산").Cells(j, 7).Value = "미달"
        End If
    Next j
End Sub

' 부서 이름의 공백을 무시하고 다시 센다.
Function 공백무시합계(dept As String) As Currency
    Dim i As Long
    Dim s As String

    For i = 4 To 19
        s = Replace(Sheets("대장").Cells(i, 2).Value, " ", "")
        If s = Replace(dept, " ", "") Then
            공백무시합계 = 공백무시합계 + Sheets("대장").Cells(i, 4).Value
        End If
    Next i
End Function

