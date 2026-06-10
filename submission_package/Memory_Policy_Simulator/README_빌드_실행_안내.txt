Memory Policy Simulator 빌드/실행 안내

1. 실행 파일
- Memory_Policy_Simulator\bin\Release\Memory_Policy_Simulator.exe

2. Visual Studio 빌드
- Memory_Policy_Simulator.sln 파일을 Visual Studio에서 연다.
- 프로젝트 대상은 .NET Framework v4.8이다.
- 빌드 PC에 .NET Framework 4.8 Developer Pack이 필요할 수 있다.

3. 구현 정책
- FIFO
- NUR (0,1 first): R/M bit class 중 (0,1)을 (1,0)보다 먼저 교체 후보로 본다.
- NUR (1,0 first): R/M bit class 중 (1,0)을 (0,1)보다 먼저 교체 후보로 본다.
- Second Chance
- LRFU-Lite

4. 사용 방법
- Policy 콤보박스에서 교체 정책을 선택한다.
- Reference String에는 한 글자 단위 페이지 참조열을 입력한다.
- #Frame에는 프레임 수를 입력한다.
- Clock Start에는 Second Chance의 초기 clock hand 위치를 1부터 시작하는 값으로 입력한다.
- R Reset에는 NUR의 reference bit를 몇 번의 참조마다 초기화할지 입력한다.
- Modified Pages에는 modified bit를 1로 둘 페이지 문자를 입력한다. 예: AD
- Run을 누르면 단계별 결과, 프레임 전이, chart, fault rate와 함께 각 알고리즘의 내부 상태가 출력된다.
