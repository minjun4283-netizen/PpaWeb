# vendor/

동봉된 서드파티 라이브러리 (둘 다 MIT 라이선스, 순수 Python, 컴파일된 바이너리
없음 — 그래서 폴더째로 아무 OS에 복사해도 그대로 동작합니다):

- `openpyxl` 3.1.5 — xlsm/xlsx 읽기. <https://pypi.org/project/openpyxl/>
- `et_xmlfile` 2.0.0 — openpyxl의 의존 패키지. <https://pypi.org/project/et-xmlfile/>

라이선스 전문은 각 `*.dist-info/LICENCE.rst`에 있습니다.

`ppa_loader.py`가 이 폴더를 자동으로 `sys.path`에 추가해서 사용하므로,
`pip install openpyxl` 없이도 (그리고 인터넷 연결 없이도) xlsm 방식이
그대로 동작합니다. `static-dashboard` 폴더를 옮길 때 `vendor/`를 통째로
같이 옮기기만 하면 됩니다.
