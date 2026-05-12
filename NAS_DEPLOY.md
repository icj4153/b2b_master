# NAS 자동 실행 설정

이 프로젝트는 Docker Compose로 NAS에서 매일 한 번 `b2b_excel.py`를 실행할 수 있습니다.

## 1. NAS 준비

- NAS에 Docker 또는 Container Manager가 설치되어 있어야 합니다.
- 프로젝트 폴더 전체를 NAS의 원하는 위치에 업로드합니다.
- NAS의 프로젝트 폴더에 `.env` 파일을 만들고 로그인/API 값을 채웁니다.

```text
B2B_USER_ID=...
B2B_USER_PW=...
NAVER_CLIENT_ID=...
NAVER_CLIENT_SECRET=...
```

## 2. 실행

NAS 터미널에서 프로젝트 폴더로 이동한 뒤 실행합니다.

```bash
cd /volume1/docker/b2b_excel
PATH=/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/local/bin/docker-compose up -d --build
```

기본 실행 시간은 매일 한국 시간 오전 9시입니다. 컨테이너 시간대는 `KST-9`로 설정되어 있습니다.

## 3. 실행 시간 변경

`docker-compose.yml`의 `B2B_RUN_AT` 값을 바꾸면 됩니다.

```yaml
B2B_RUN_AT: "08:30"
```

변경 후 다시 적용합니다.

```bash
cd /volume1/docker/b2b_excel
PATH=/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/local/bin/docker-compose up -d
```

## 4. 결과 파일 위치

통합 엑셀 결과는 NAS 프로젝트 폴더의 `output` 폴더에 저장됩니다.
임시 다운로드 파일은 `b2b_downloads` 폴더를 사용하고, 스크립트 종료 시 정리됩니다.

## 5. 수동으로 바로 한 번 실행

스케줄을 기다리지 않고 바로 테스트하려면 다음 명령을 실행합니다.

```bash
cd /volume1/docker/b2b_excel
PATH=/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/local/bin/docker-compose run --rm b2b-excel python b2b_excel.py
```

## 6. 로그 확인

```bash
cd /volume1/docker/b2b_excel
PATH=/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/local/bin/docker-compose logs -f b2b-excel
```

## 7. GitHub에서 최신 코드 반영

NAS 폴더가 GitHub clone 기반이면 아래 명령으로 최신 코드를 받고 컨테이너를 다시 빌드합니다.

```bash
cd /volume1/docker/b2b_excel
sh nas_update.sh
```

## 8. 외부 접속

Streamlit 웹서비스는 컨테이너에서 8501 포트로 실행됩니다. 외부 공개 시에는 공유기에서 NAS로 포트포워딩하거나 Synology DDNS와 Reverse Proxy를 사용합니다.

## 9. 웹서비스 접속

Streamlit 웹서비스는 NAS의 8501 포트로 실행됩니다.

```text
http://192.168.50.101:8501
```
