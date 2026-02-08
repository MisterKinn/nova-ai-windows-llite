# Nova AI Lite (경량화 버전)

Nova AI Lite는 기존 프로그램의 핵심 기능( AI 연결, 한글(HWP) 자동화, 수식 삽입 )만 남긴
최소 구성 버전입니다. GUI 없이 CLI로 동작하며, Windows 환경을 기본 대상으로 합니다.

## 구성 요약
- `ai_client.py`: Gemini 연결 (단일 모델)
- `hwp_controller.py`: HWP 연결/텍스트 입력
- `equation.py`: 수식 객체 삽입 (HwpEqn 문법)
- `script_runner.py`: 최소 샌드박스 실행기
- `app.py`: CLI 엔트리포인트

## 설치
```bash
pip install -r requirements.txt
```

## 사용 예시
```bash
python app.py detect
python app.py insert-text "안녕하세요"
python app.py insert-equation "a over b"
python app.py insert-equation "x^2 + y^2 = z^2" --latex
python app.py insert-latex-equation "x^2 + y^2 = z^2"
python app.py run-script --file my_script.py
python app.py ai-generate "문제를 번호 붙여 입력해줘" --output out.py
python app.py ai-run "x^2 + y^2 = z^2 를 수식으로 입력"
```

## GUI 실행
```bash
python gui_app.py
```

## GUI 동작
- 사진을 업로드하면 AI가 코드를 생성
- 생성된 코드를 로그에 표시
- 동시에 HWP에 자동 입력

## 환경변수
- `GEMINI_API_KEY`: AI 사용 시 필수
- `NOVA_AI_MODEL`: 기본 모델 지정 (예: `gemini-3-flash-preview`)

## 참고
Nova AI Lite는 기본적으로 **HWP 수식 문법(HwpEqn)** 을 사용합니다.
또한 `node_eqn/hwp_eqn_cli.js`가 존재하면 LaTeX → HwpEqn 변환도 지원합니다.
