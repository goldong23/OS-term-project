Memory Policy Simulator Build/Run Guide

1. Executable
- Memory_Policy_Simulator\bin\Release\Memory_Policy_Simulator.exe

2. Visual Studio build
- Open Memory_Policy_Simulator.sln in Visual Studio.
- Target framework: .NET Framework 4.8.
- A PC without the .NET Framework 4.8 Developer Pack may fail to build the project file.

3. Implemented policies
- FIFO
- NUR (0,1 first): selects victims by R/M bit class with (0,1) before (1,0).
- NUR (1,0 first): selects victims by R/M bit class with (1,0) before (0,1).
- Second Chance: uses reference bits and a circular clock hand.
- WSClock-Lite: uses a clock hand, R/M bits, last-use age, and an age threshold.

4. Options
- Policy: selects the page replacement policy.
- Reference String: treats each character as one page reference.
- #Frame: sets the number of page frames.
- Clock: resets R bits every N steps. It is used by NUR and Second Chance. In WSClock-Lite, the same value is also used as the age threshold.
- Modified Pages: marks listed page characters as M=1. It is used by NUR and WSClock-Lite.
- The UI only displays the options used by the currently selected policy.

5. Run
- Press Run to print each step, frame state, victim page, policy state, chart, and page fault rate.
