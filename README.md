# 📝 HWP/HWPX 병합기 (pyhwpx 기반)

한글(HWP/HWPX) 파일들을 하나로 병합하고, PDF로도 저장할 수 있는 **GUI 프로그램**입니다.  
Windows + 한글(HWP) 설치 환경에서 동작하며, `pyhwpx`와 `pywin32` COM 연동을 사용합니다.

---

## 📂 프로젝트 구조
```text
P3. SUM/
 ├─ build/                # PyInstaller 빌드 중간 산출물
 ├─ dist/                 # 최종 실행파일(.exe) 생성 경로
 │   └─ merge_hwp.exe     # 빌드된 실행파일
 ├─ hwp_icon.ico          # 프로그램 아이콘 (PyInstaller 빌드 시 사용)
 ├─ hwp_icon.png          # 원본 PNG 아이콘
 ├─ merge_hwp.py          # 메인 프로그램 (Python 코드)
 ├─ merge_hwp.spec        # PyInstaller 스펙 파일
 ├─ sum.py, sum2.py       # 테스트/샘플 스크립트
```
---

## 🚀 실행 방법

### 1) 직접 실행 (개발자용)
```bash
pip install pyhwpx pywin32
python merge_hwp.py
```
- GUI 창이 뜨면 병합할 `.hwp / .hwpx` 파일들을 선택합니다.  
- 병합 후 결과 파일은 `.hwp`와 (옵션 선택 시) `.pdf`로 저장됩니다.  

### 2) 실행파일 사용 (배포용)
- `dist/merge_hwp.exe` 파일만 있으면 실행 가능합니다.  
- **Python 설치 불필요**  
- 단, 실행 PC에는 **한글(HWP 프로그램)** 이 반드시 설치되어 있어야 합니다.  

---

## 🔨 빌드 방법 (개발자)

PyInstaller로 단일 실행파일 생성:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=hwp_icon.ico merge_hwp.py
```
- 빌드 후 `dist/merge_hwp.exe` 가 생성됩니다.

> 참고: 경로에 공백이 있을 경우 아이콘 옵션은 따옴표로 감싸세요.  
> `--icon "hwp_icon.ico"`

---

## ⚙️ 주요 기능
- 여러 HWP/HWPX 파일을 순서대로 병합  
- 첫 파일을 기반 문서로 사용 → **공백 첫 페이지 방지**  
- 서식 유지 옵션 (쪽 구역, 글자 모양, 문단 모양, 스타일)  
- 병합 결과를 **PDF로도 자동 저장 가능**  
- 프로그램 하단 제작자 표시:  
  `마산구암고등학교 미래교육지원부 (2025)`  

---

## 📌 요구 사항
- Windows 환경  
- 한글(HWP 프로그램) 설치 필수  
- (개발/빌드 시) Python 3.9+ 권장  

---

## 👤 제작자
**Teacher Cho**  
마산구암고등학교 미래교육지원부 (2025)
