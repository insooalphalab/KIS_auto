"""
시크릿/개인 설정 로더.

프로젝트 루트의 .env 파일을 읽어 환경변수로 노출한다 (외부 패키지 불필요).
.env 는 git 에 올리지 않는다 (.gitignore 참고). 템플릿은 .env.example.

이미 OS 환경변수로 설정된 값이 있으면 .env 값보다 우선한다.
"""

import os

_ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")


def _load_dotenv(path: str) -> None:
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)   # base64 값에 '=' 가 들어 있어 첫 '=' 에서만 분리
            key, value = key.strip(), value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
                value = value[1:-1]
            os.environ.setdefault(key, value)


_load_dotenv(_ENV_PATH)


def require(name: str) -> str:
    """필수 설정값을 반환. 없으면 설정 방법을 안내하며 즉시 중단한다."""
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(
            f"설정값 {name} 이(가) 없습니다. .env.example 을 .env 로 복사한 뒤 값을 채워주세요."
        )
    return value
