# 엑셀 서식 도우미 MVP

엑셀 파일과 자연어 요청을 받아 간단한 결과 통합문서(`.xlsx`)를 만드는 Windows 프로그램입니다.

현재 지원 파일은 `.xlsx`, `.xlsm`, `.csv`이며 결과는 `.xlsx`로 저장됩니다.

현재 지원하는 요청:

- 그룹별 합계: `이름을 기준으로 금액을 합산하고 큰 순서대로 보여줘`
- 조건 필터: `상태가 완료인 행만 보여줘`
- 정렬: `금액을 큰 순서로 정렬해줘`
- 중복 찾기: `이름 중복 항목을 찾아줘`
- 그룹별 평균·개수와 최댓값·최솟값·평균 통계
- 같은 값, 다른 값, 포함, 이상·초과·이하·미만 조건 필터

## 실행

```powershell
& 'C:\Users\ghddm\.pyenv\pyenv-win\versions\3.12.4\python.exe' app.py
```

## 테스트

```powershell
& 'C:\Users\ghddm\.pyenv\pyenv-win\versions\3.12.4\python.exe' -m unittest discover -s tests -v
```

## EXE 빌드

```powershell
.\build.ps1
```

빌드 결과는 `dist\너무 엑셀 팡션 사용하지 마세요.exe`입니다. 다음 빌드 시 같은 이름의 이전 실행 파일을 덮어씁니다.
