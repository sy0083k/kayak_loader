' ================================================================================
' kayak_vba.bas  —  카약 2대 적재 계획 유틸리티 매크로
' ================================================================================
'
' ※ 적재_계획 F/G열은 이제 수식으로 자동 반영됩니다.
'    체크박스 클릭 → H/I/J/K TRUE/FALSE 변경
'    → 적재_계획 F열 =IF(장비_DB!H{n},"선수",IF(장비_DB!I{n},"선미",""))
'    → 적재_계획 G열 =IF(장비_DB!J{n},"선수",IF(장비_DB!K{n},"선미",""))
'    VBA 없이도 즉시 반영됩니다.
'
' [선택적 VBA 설치 — 전체 초기화 버튼이 필요할 때만]
'   1. Alt+F11 → VBA 에디터
'   2. 삽입 → 모듈 → Module1에 아래 코드 붙여넣기
' ================================================================================


' ── 표준 모듈 (Module1) ──────────────────────────────────────────────────────

' 전체 초기화: 체크박스(H:K)를 모두 FALSE로 → F/G 수식이 자동으로 빈칸이 됨
Sub ClearAllPlan()
    If MsgBox("모든 배치를 초기화합니다. 계속하시겠습니까?", vbYesNo) <> vbYes Then Exit Sub
    ThisWorkbook.Sheets("장비_DB").Range("H2:K51").Value = False
    MsgBox "초기화 완료 — 적재_계획이 자동 갱신됩니다.", vbInformation
End Sub


' 선수/선미 중복 방지 (같은 카약에 선수+선미 동시 체크 불가)
' → 체크박스에 직접 매크로 연결 방식으로 사용
'   개발 도구 → 삽입 → 양식 컨트롤 체크박스 우클릭 → "매크로 지정"
'   각 행/열에 맞는 Sub를 지정하거나, 아래 범용 Sub 하나를 모든 체크박스에 지정
Sub MutualExclude()
    ' 현재 선택된 체크박스의 LinkedCell을 기준으로 상호 배제
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("장비_DB")
    Dim cb As Object
    Set cb = ws.Shapes(Application.Caller)
    Dim lnkCell As Range
    Set lnkCell = ws.Range(cb.ControlFormat.LinkedCell)
    If lnkCell.Value = False Then Exit Sub   ' 체크 해제 시 무시
    Dim r As Long, c As Long
    r = lnkCell.Row
    c = lnkCell.Column
    Select Case c
        Case 8:  ws.Cells(r, 9).Value = False   ' K1선수 체크 → K1선미 해제
        Case 9:  ws.Cells(r, 8).Value = False   ' K1선미 체크 → K1선수 해제
        Case 10: ws.Cells(r, 11).Value = False  ' K2선수 체크 → K2선미 해제
        Case 11: ws.Cells(r, 10).Value = False  ' K2선미 체크 → K2선수 해제
    End Select
End Sub
