using System;

namespace Memory_Policy_Simulator
{
    class AlgorithmSmokeTest
    {
        static int Run(string referenceString, int frameSize, Core.ReplacementPolicy policy, int clockStart, int resetInterval, string modifiedPages)
        {
            Core core = new Core(frameSize, policy, clockStart, resetInterval, modifiedPages);

            for (int i = 0; i < referenceString.Length; i++)
            {
                core.Operate(referenceString[i], i, referenceString);
            }

            return core.fault;
        }

        static void Expect(string name, int actual, int expected)
        {
            Console.WriteLine(name + ": " + actual + " page faults");
            if (actual != expected)
            {
                Environment.ExitCode = 1;
            }
        }

        static void Main()
        {
            string referenceString = "123412512345";

            Expect("FIFO/4", Run(referenceString, 4, Core.ReplacementPolicy.FIFO, 1, 4, "14"), 10);
            Expect("NUR(0,1 first)/4", Run(referenceString, 4, Core.ReplacementPolicy.NUR_01_First, 1, 4, "14"), 7);
            Expect("NUR(1,0 first)/4", Run(referenceString, 4, Core.ReplacementPolicy.NUR_10_First, 1, 4, "14"), 7);
            Expect("SecondChance/4 clock=2", Run(referenceString, 4, Core.ReplacementPolicy.SecondChance, 2, 4, "14"), 7);
            Expect("LRFU-Lite/4", Run(referenceString, 4, Core.ReplacementPolicy.LRFULite, 1, 4, "14"), 8);
            Console.WriteLine(Environment.ExitCode == 0 ? "Smoke test passed." : "Smoke test failed.");
        }
    }
}
