Memory Policy Simulator 빌드/실행 안내

1. 실행 파일
- Memory_Policy_Simulator\bin\Release\Memory_Policy_Simulator.exe

2. Visual Studio 빌드
- Memory_Policy_Simulator.sln 파일을 Visual Studio에서 연다.
- 프로젝트 대상은 .NET Framework v4.8이다.
- 빌드 PC에 .NET Framework 4.8 Developer Pack이 필요할 수 있다.

3. 구현 정책
- FIFO
- Optimal
- LRU
- Second Chance
- LRFU-Lite

4. 사용 방법
- Policy 콤보박스에서 교체 정책을 선택한다.
- Reference String에는 한 글자 단위 페이지 참조열을 입력한다.
- #Frame에는 프레임 수를 입력한다.
- Run을 누르면 단계별 결과, 프레임 전이, chart, fault rate가 출력된다.
