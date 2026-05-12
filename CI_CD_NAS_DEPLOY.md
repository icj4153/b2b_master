# GitHub Actions NAS CI/CD

`main` 브랜치에 push되면 GitHub Actions가 NAS에 SSH 접속해 `/volume1/docker/b2b_excel/nas_update.sh`를 실행합니다. NAS는 기존 clone 폴더에서 최신 코드를 pull하고 Docker Compose로 크롤러와 웹 컨테이너를 재빌드/재시작합니다.

## 구조

- GitHub Actions workflow: `.github/workflows/deploy-nas.yml`
- NAS 배포 진입점: `/volume1/docker/b2b_excel/nas_update.sh`
- NAS 앱 경로: `/volume1/docker/b2b_excel`
- 외부 웹 주소: `https://icj7297.synology.me/`

## GitHub Secrets

Repository `Settings > Secrets and variables > Actions > New repository secret`에 아래 값을 등록합니다.

| Secret | Value |
| --- | --- |
| `NAS_HOST` | `icj7297.synology.me` |
| `NAS_PORT` | `2022` |
| `NAS_USER` | `joon_admin` |
| `NAS_SSH_PRIVATE_KEY` | GitHub Actions 전용 SSH private key 전체 내용 |
| `NAS_KNOWN_HOSTS` | `ssh-keyscan -p 2022 icj7297.synology.me` 결과 |

## SSH 키 등록

GitHub Actions 전용 SSH key pair를 만들고 public key를 NAS의 `joon_admin` 계정에 등록합니다. private key는 GitHub Secret `NAS_SSH_PRIVATE_KEY`에만 저장하고 저장소에는 커밋하지 않습니다.

NAS에 public key가 등록되면 GitHub Actions는 아래 명령을 실행할 수 있어야 합니다.

```bash
ssh -p 2022 joon_admin@icj7297.synology.me "cd /volume1/docker/b2b_excel && sh nas_update.sh"
```

## 네트워크 조건

공유기 포트포워딩이 필요합니다.

```text
외부 TCP 2022 -> NAS 192.168.50.101:2022
```

웹 접속용 포트포워딩은 별도로 유지합니다.

```text
외부 TCP 443 -> NAS 192.168.50.101:443
외부 TCP 80  -> NAS 192.168.50.101:80
```

## 배포 확인

GitHub Actions 실행 후 NAS에서 확인합니다.

```bash
cd /volume1/docker/b2b_excel
git log -1 --oneline
/usr/local/bin/docker ps --filter name=b2b-excel
```

웹은 아래 주소에서 확인합니다.

```text
https://icj7297.synology.me/
```

크롤러 스케줄 로그는 아래 명령으로 확인합니다.

```bash
cd /volume1/docker/b2b_excel
/usr/local/bin/docker-compose logs --tail=30 b2b-excel
```

## 장애 확인

- Secret 누락: Actions의 `Validate required secrets` 단계가 실패합니다.
- SSH 실패: `NAS_HOST`, `NAS_PORT`, `NAS_USER`, `NAS_SSH_PRIVATE_KEY`, `NAS_KNOWN_HOSTS` 값을 확인합니다.
- 배포 실패: NAS에서 직접 `cd /volume1/docker/b2b_excel && sh nas_update.sh`를 실행해 원인을 확인합니다.
- 웹 접속 실패: Reverse Proxy, 443 포트포워딩, Streamlit WebSocket 설정을 확인합니다.
