# legacy — 통합 전 원본 보존 폴더

아래 두 파일은 `monitor/hanyang_monitor.py` 통합 프로그램으로 대체되었지만,
**원본 그대로 보존**하기 위해 이 폴더에 남겨 둡니다. (삭제하지 않음)

| 파일 | 설명 | 통합본 위치 |
|------|------|-------------|
| `Client.py` | 백그라운드 모니터링 에이전트 (구버전) | `monitor/hanyang_monitor.py` 의 `MonitoringAgent` |
| `viewer_v4.py` | 관제 대시보드 (구버전) | `monitor/hanyang_monitor.py` 의 `MonitorViewer` |

- 실제 운영/배포에는 `monitor/hanyang_monitor.py` 를 사용하세요.
- 이 파일들은 참고 및 롤백 용도로만 유지됩니다.
