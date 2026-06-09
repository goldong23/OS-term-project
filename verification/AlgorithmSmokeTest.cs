using System;

namespace Memory_Policy_Simulator
{
    class AlgorithmSmokeTest
    {
        static int Run(string referenceString, int frameSize, Core.ReplacementPolicy policy)
        {
            Core core = new Core(frameSize, policy);

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

            Expect("FIFO/4", Run(referenceString, 4, Core.ReplacementPolicy.FIFO), 10);
            Expect("Optimal/4", Run(referenceString, 4, Core.ReplacementPolicy.Optimal), 6);
            Expect("LRU/4", Run(referenceString, 4, Core.ReplacementPolicy.LRU), 8);
            Console.WriteLine(Environment.ExitCode == 0 ? "Smoke test passed." : "Smoke test failed.");
        }
    }
}
