# KakaoTalk UI Automation Script

Windows용 카카오톡 PC 클라이언트를 자동으로 제어해 특정 채팅방을 열고 메시지를 전송한 뒤, 수신 메시지를 모니터링하는 파이썬 스크립트를 제공합니다. 주요 동작은 Win32 API, `uiautomation`, `pywinauto`를 조합해 구현되었습니다.

## 주요 기능
- 지정한 채팅방 검색 및 진입 (`open_chatroom`).
- 입력창을 탐색한 뒤 실제 키보드 타이핑을 흉내 내거나 여러 백업 입력 방식을 사용해 메시지 전송 (`send_message_and_verify`).
- 최근 대화 내용을 클립보드로 읽어 전송 성공 여부 검증 및 실시간 수신 모니터링 (`get_chat_text`).
- 디버깅을 위한 상세 로그 출력 옵션(`DEBUG`).

## 환경 요구 사항
- **운영체제**: Windows (Win32 API 사용)
- **Python**: 3.x (스크립트는 Python 3 기준 작성)
- **필수 패키지**: `pywin32`, `uiautomation`, `pywinauto`

```bash
pip install pywin32 uiautomation pywinauto
```

> 스크립트는 Win32 메시지와 `SendInput`을 직접 사용하므로 Windows 외 OS에서는 동작하지 않습니다.

## 설정 방법
`uiautomation_kakao2.py` 상단의 사용자 설정 값을 수정해 원하는 대상과 동작을 정의합니다.

```python
CHATROOM_NAME = "차희상"          # 열고자 하는 채팅방 이름
MESSAGE_TEXT  = "안녕하세요..."   # 전송할 메시지 내용
DEBUG = True                      # True면 상세 로그 출력
VERIFY_TAIL_LINES = 30            # 전송 후 하단 몇 줄로 검증할지
OPEN_SCAN_LIMIT = 500             # 채팅방 목록에서 스캔 최대 횟수
PAGEDOWN_EVERY = 30               # 목록 이동 중 PageDown 주기
BRING_TO_FRONT = True             # 카카오톡 창을 전경에 강제 이동할지
```

## 실행 방법
1. 카카오톡 PC 클라이언트를 실행한 상태에서 터미널(명령 프롬프트 또는 PowerShell)을 엽니다.
2. 스크립트가 있는 디렉터리로 이동합니다.
3. 아래 명령으로 실행합니다.

```bash
python uiautomation_kakao2.py
```

실행되면 스크립트가 지정한 채팅방을 열고 메시지를 전송한 뒤, 전송 성공 여부를 확인합니다. 이후 새 메시지를 실시간으로 모니터링하며 변경 사항을 출력합니다.

## 동작 개요
- **채팅방 열기**: 카카오톡 메인 창과 내부 컨트롤(EVA_ChildWindow/EVA_Window 등)을 찾아 검색창에 방 이름을 입력하고, 리스트 포커스를 이동해 엔터로 진입합니다.
- **입력창 찾기**: 대화 리스트 하단 근처의 `RichEdit`/`Edit` 컨트롤을 우선 탐색하고, `WM_SETTEXT/GETTEXT` probe로 실제 입력이 가능한지를 검사합니다.
- **메시지 전송**: `SendInput`을 이용한 실제 타이핑 → UIA SendKeys → 클립보드 붙여넣기 → `WM_SETTEXT` 순으로 여러 입력 경로를 시도하며, 각 단계마다 엔터 전송 변형을 반복합니다.
- **검증/모니터링**: 대화 리스트 내용을 클립보드로 복사해 전송 성공 여부를 확인하고, 이후 주기적으로 차이를 감지해 새 메시지를 출력합니다.

## 주의사항
- 카카오톡 UI나 클래스명이 변경되면 컨트롤 탐색 로직이 실패할 수 있습니다. 필요 시 `_find_child_by_class_recursive`, `_find_input_edit_win32`의 우선순위나 필터 기준을 조정하세요.
- 자동 입력 중 다른 앱이 포커스를 빼앗거나 카카오톡 창이 최소화되어 있으면 동작이 불안정해질 수 있습니다. `BRING_TO_FRONT`를 `True`로 유지하는 것을 권장합니다.
- 실제 입력을 시뮬레이션하므로 테스트 시 주의하십시오. 기본 메시지와 대상 채팅방을 변경해 사용하세요.

## 라이선스
별도의 라이선스 정보가 없으므로 필요 시 저장소 소유자에게 문의하세요.
