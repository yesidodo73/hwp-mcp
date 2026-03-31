# hwp-mcp 프로젝트 구조 설명

## 개요

hwp-mcp는 한글 워드 프로세서(HWP)를 Claude와 같은 AI 모델이 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버입니다. 이 프로젝트는 한글 문서를 자동으로 생성, 편집, 조작하는 기능을 Claude에게 제공합니다.

## 작동 방식

1. **Claude 호출 과정**:
   - Claude 데스크톱 애플리케이션에서 hwp-mcp를 호출하면 `경로/hwp-mcp/hwp_mcp_stdio_server.py` 스크립트가 실행됩니다.
   - 이 스크립트는 표준 입출력(stdio)을 통해 Claude와 통신하는 FastMCP 서버를 생성합니다.
   - 서버는 HWP를 조작하기 위한 여러 도구(tool)들을 등록합니다.

2. **기능 실행 과정**:
   - Claude가 HWP 관련 기능을 요청하면, 등록된 도구가 호출됩니다.
   - 도구는 `HwpController`를 사용하여 Windows COM 자동화를 통해 HWP 프로그램과 상호작용합니다.
   - `win32com` 라이브러리를 통해 HWP에 명령을 전송합니다.

## 주요 컴포넌트

1. **hwp_mcp_stdio_server.py**:
   - 프로젝트의 진입점(entry point)
   - FastMCP 서버를 초기화하고 도구들을 등록
   - stdio를 통해 Claude와 통신하는 인터페이스 제공

2. **src/tools/hwp_controller.py**:
   - HWP와 상호작용하는 핵심 컨트롤러 클래스
   - win32com을 이용하여 한글 프로그램을 자동화
   - 문서 생성, 열기, 저장, 텍스트 삽입, 테이블 생성 등의 기능 제공

3. **security_module/**:
   - 파일 경로 체크 보안 경고창을 우회하기 위한 모듈 참조
   - 실제 파일은 확인되지 않음

## 의존성

1. **Python 패키지 의존성** (requirements.txt):
   - pywin32 (HWP 자동화를 위한 핵심 라이브러리)
   - comtypes
   - pytest
   - pytest-cov
   - mcp (`mcp.server.fastmcp` 제공)

2. **외부 의존성**:
   - `mcp` Python 패키지: 이 프로젝트의 핵심 외부 의존성
     - `FastMCP`를 `mcp.server.fastmcp`에서 임포트
     - Claude가 HWP 프로그램과 통신할 수 있게 해주는 Model Context Protocol 구현체

## 프로젝트 삭제 시 고려사항

이 프로젝트를 삭제하고자 할 경우 다음 사항을 고려해야 합니다:

1. 이 코드는 Claude와 HWP 사이의 브릿지 역할을 합니다.
2. 외부 패키지인 `mcp`에 의존성이 있으며, 이는 별도로 설치되어 있을 가능성이 높습니다.
3. Claude 데스크톱 설정 파일에 이 프로젝트에 대한 참조가 있습니다.

프로젝트를 삭제할 경우 Claude는 더 이상 HWP를 제어할 수 없게 됩니다. 이것이 의도한 바라면, Claude 설정에서 hwp mcp 항목을 제거하는 것도 함께 진행해야 합니다.

## 설정 파일 예시

```json
{
  "mcpServers": {
    "hwp": {
      "command": "python",
      "args": ["경로/hwp-mcp/hwp_mcp_stdio_server.py"]
    }
  }
}
```

위 설정에서 "hwp" 항목을 제거하면 Claude가 더 이상 HWP MCP를 호출하지 않습니다. 
