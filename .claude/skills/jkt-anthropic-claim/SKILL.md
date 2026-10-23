---
name: jkt-anthropic-claim
description: 진정호 팀장(JKT Electric)의 Anthropic(Claude Max) 구독료 월별 ใบเบิก(PURCHASE REQUEST) 정산 묶음을 만든다. inbox/YYMM 의 Anthropic Invoice·Receipt PDF(+사내 PR PDF, KBank 결제 문자)를 읽어 ใบเบิก 엑셀 작성·서명 → PDF 변환 → ใบเบิก/Invoice/Receipt/서명PR/문자 순 합본까지 한 번에 처리한다. 사용자가 "앤트로픽 정산", "Anthropic 정산", "클로드 구독료 정산", "Claude 구독 ใบเบิก", "AI 구독료 청구", "anthropic claim" 등을 말하거나 Anthropic Invoice-TNXTMSWM-*.pdf / Receipt-*.pdf 를 올리며 정산·ใบเบิก를 요청하면 반드시 이 스킬을 사용하라.
---

# Anthropic 구독료 월별 ใบเบิก 정산

매달 같은 패턴. 결정 규칙은 아래에 고정돼 있고, 반복 작업은 `scripts/build_claim.py` 가 한다.
**원본(템플릿·inbox 서류)은 절대 덮어쓰지 않는다.** 결과물은 전부 `out/` 에 새로 만든다.

## 폴더 (`anthropic-claim/`, 저장소 루트 기준)

```
anthropic-claim/
  template/แบบฟอร์มใบเบิก.xlsx     # JKT ใบขอซื้อ 원본 양식 — 수정 금지
  template/sig_chin.png            # 진정호 서명 (195×93, 투명 배경)
  inbox/YYMM/                      # 그달 원본 서류
    Invoice-TNXTMSWM-00NN.pdf      # Anthropic 인보이스 (USD)
    Receipt-XXXX-XXXX-XXXX.pdf     # Anthropic 영수증 (결제일·카드 끝 4자리)
    PR-YYMMDDNN.pdf                # 사내 PR (ใบขอซื้อวัสดุ) — 바트 금액 근거
    sms.png                        # (선택) KBank 카드결제 문자 캡처
  out/                             # 결과물
  history.md                       # 월별 이력 (끝나면 한 줄 추가)
```

폴더나 파일이 없으면 만들고, 무엇을 어디에 넣어야 하는지 팀장에게 알려준다.
업로드된 파일명이 다르면(`Invoice-…_1.pdf` 등) inbox 에 위 규칙 이름으로 **복사**한다.
`template/แบบฟอร์มใบเบิก.xlsx` 가 없으면 스크립트가 `out/` 의 지난 완성본을 바탕으로 쓴다
(10행·Vat/WHT·서명 이미지는 매번 새로 쓰므로 안전). 그것도 없으면 `--base` 로 직전 ใบเบิก 를 준다.

Invoice·Receipt 받는 곳: claude.ai → 왼쪽 아래 이름 → Settings → Billing → Invoices → 해당 건 View,
또는 메일 제목 "Your receipt from Anthropic".

## 바트 금액 (F10) — 이 순서만 허용

1. PR PDF 의 Price 칸 (사내 PR 이 있으면 이게 기준)
2. 없으면 KBank 문자 `ชำระ X,XXX.XXบ บัตร x-NNNN@ANTHROPIC* CLAUDE S` 의 금액
3. 둘 다 없으면 **F10 을 비우고 팀장에게 묻는다. 환율로 추정해서 넣지 않는다.**
   (`--thb` 없이 돌리면 스크립트가 F10 을 비우고 파일명에 `_DRAFT` 를 붙인다)

PR·문자는 Claude 가 직접 읽어서 금액을 확인한 뒤 `--thb` 로 넘긴다.
PR 이 있으면 스크립트가 PR 본문에 그 금액이 있는지 한 번 더 대조한다.

## 실행

```bash
S=.claude/skills/jkt-anthropic-claim/scripts/build_claim.py

# 1) PR 서명 칸 먼저 확인 (Review 두 곳이 잡히는지) — jkt-sign 스크립트
SIGN=$(find ~/.claude/skills -path '*jkt-sign/scripts/sign.py' | head -1)
python3 "$SIGN" --pdf anthropic-claim/inbox/YYMM/PR-….pdf --list-anchors

# 2) 정산 묶음 생성
python3 $S --month YYMM --thb 3436.81 --pr anthropic-claim/inbox/YYMM/PR-26092803.pdf
python3 $S --month YYMM --thb 3436.81 --sms anthropic-claim/inbox/YYMM/sms.png --mask x0,y0,x1,y1
python3 $S --month YYMM                                   # 금액 근거 없음 → _DRAFT
```

옵션: `--pr-signed`(이미 서명된 PR), `--pr-no`(PR 번호 직접), `--base`(양식 경로), `--date YYMMDD`(작성일, 기본 오늘).

## 스크립트가 하는 일 (규칙 고정 — 바꾸려면 팀장 확인)

1. **추출**: Invoice 번호·발행일·플랜(5x/20x)·사용 기간·USD 합계·차감 라인
   (업그레이드 시 `Unused time on Max plan - 5x …` 크레딧). 영수증의 Invoice 번호·금액과 불일치면 중단.
2. **ใบเบิก 엑셀** `out/YYMMDD_ใบเบิก_Anthropic_JEONGHO CHIN.xlsx`
   - C6 `JEONGHO CHIN`, C7 `JKT` 유지, H6 `=TODAY()`
   - B10 1 / C10 `ค่าบริการ AI รายเดือน Claude Max plan {5x|20x} (Anthropic)`
     (+업그레이드면 ` อัปเกรดจาก 5x รอบ DD/MM/YYYY-DD/MM/YYYY`)
   - D10 1 / E10 `เดือน` / F10 바트 / G10 `=D10*F10`
   - H10 `Invoice TNXTMSWM-00NN USD {usd} = THB {thb} ตามยอดเรียกเก็บบัตร {영수증 결제일} / อ้างอิง {PR번호}`
     업그레이드 달은 USD 뒤에 `(Max 20x ราคา 200.00 หักเครดิต Max 5x ที่เหลือ 16.73)` 를 넣는다.
     (`20x 200.00` 처럼 숫자끼리 붙이면 폰트가 `20×` 로 찍으므로 사이에 `ราคา` 를 둔다)
   - C10·H10 줄바꿈 + 왼쪽 정렬, H10 11pt, 10행 높이 72
   - Vat 7%·WHT 값 칸(G18·G19) 비움 — VAT 포함 실지출 정산
   - 서명: sig_chin.png 를 OneCellAnchor col=2(C) colOff=40px, row=21(22행) rowOff=4px,
     높이 52px 비율 유지 → "Request by" 점선 바로 위
   - 로고: openpyxl 이 저장 시 crop(srcRect)을 버리므로 crop 을 이미지에 구워서 보존한다
3. **PR 서명**: 요청부서·자재부서의 **Review 칸 두 곳에만**
   (`jkt-sign --anchor Review --occurrence 0 --height 16`). Charger·Manager 칸은 비운다.
   원본 PDF 는 재생성하지 않고 오버레이 merge. 결과 `out/PR-…_signed.pdf`.
   ※ jkt-sign 의 PR 기본값(Manager#1)과 다르다 — 이 정산은 팀장 지시로 Review 칸이다.
4. **문자 캡처**(있을 때만): Anthropic 외 개인 결제 말풍선은 Claude 가 이미지를 보고 영역을 정해
   `--mask x0,y0,x1,y1` 로 넘긴다 → 글자 없이 회색 블록(198,202,206). "MASKED" 같은 문구 금지.
5. **합본** `out/YYMMDD_ใบเบิก_Anthropic_JEONGHO CHIN_전체.pdf`
   순서 = ใบเบิก → Invoice → Receipt → 서명한 PR (→ 문자 캡처는 있을 때만 맨 뒤)
6. **검증**: 이미지 2개(로고+서명), G20 수식 유지, PDF 1쪽·수식 오류 0,
   Grand Total = 바트 금액. 1쪽을 `out/…_p1_check.png` 로 렌더한다.

## Claude 가 직접 할 일

- 스크립트 끝나면 `_p1_check.png` 를 **눈으로 열어** 서명 위치(Request by 점선 위)와
  H10 줄바꿈(잘림 없음)을 확인한다. PR 이 있으면 서명본도 렌더해서 Review 두 칸만 바뀌었는지 본다.
- `history.md` 에 한 줄 추가: `| YYYY-MM | Invoice | USD | THB | 비고(플랜, PR번호) |`
- 보고는 짧게: **금액(USD/THB)·첨부 순서·비어 있는 칸**만.
