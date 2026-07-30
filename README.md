# 서적부 POS (Book Store POS)

사업자 등록 없이 운영하는 서점용 웹 포스기입니다. 일반 포스기를 쓸 수 없는 환경에서
버튼 클릭만으로 **재고 차감·매출 집계**가 자동으로 되고, 여러 사람이 **실시간으로 함께** 사용할 수 있습니다.

## 기능
- 판매(계좌이체/현금), 실수령액·차액 기록 (실제 금액과 장부 불일치 대응)
- 할인 (원 / % 전환)
- 반품 (부분 반품 지원)
- 재고 관리 (입고/수정/삭제, 결제수단별 판매 집계)
- 매출 내역 + CSV 내보내기
- Firebase 실시간 동기화 + 로그인(가게 공용 계정)

## 기술
- 단일 `index.html` (프레임워크 없음)
- Firebase Authentication + Firestore (실시간, 오프라인 지원)

## 실행
정적 웹 호스팅(GitHub Pages 등)에 올려서 사용합니다.
로컬 테스트 시에도 `file://` 이 아닌 웹서버(`http://localhost:...`)로 열어야 합니다.

## 설정
- `firebase-config.js` : Firebase 웹 설정 (공개돼도 되는 값)
- `firestore.rules` : 보안 규칙 (가게 계정만 접근 허용) — Firebase 콘솔에 게시
- `SETUP-firebase.md` : Firebase 초기 설정 안내
