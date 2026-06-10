using System;

namespace Memory_Policy_Simulator
{
    class AnalysisData
    {
        static void Run(string workloadName, string referenceString, int frameSize, Core.ReplacementPolicy policy, int clockStart, int resetInterval, string modifiedPages)
        {
            Core core = new Core(frameSize, policy, clockStart, resetInterval, modifiedPages);

            for (int i = 0; i < referenceString.Length; i++)
            {
                core.Operate(referenceString[i], i, referenceString);
            }

            double faultRate = Math.Round((double)core.fault / referenceString.Length * 100.0, 2);
            Console.WriteLine(workloadName + "," + referenceString + "," + frameSize + "," + policy + "," + clockStart + "," + resetInterval + "," + modifiedPages + "," + core.hit + "," + core.fault + "," + core.migration + "," + faultRate + "," + core.GetEstimatedPageFaultDelay());
        }

        static void Main()
        {
            string[] names = { "Textbook", "NURPriority", "ClockSensitive", "FrequencyBias", "LocalityShift" };
            string[] refs = { "123412512345", "ABCDABEFABGHABCD", "ABCDEABCDA", "AAAABCAAADEFAAA", "ABCABCABCDEFDEFABC" };
            int[] frames = { 3, 4, 5 };
            Core.ReplacementPolicy[] policies =
            {
                Core.ReplacementPolicy.FIFO,
                Core.ReplacementPolicy.NUR_01_First,
                Core.ReplacementPolicy.NUR_10_First,
                Core.ReplacementPolicy.SecondChance,
                Core.ReplacementPolicy.LRFULite
            };

            Console.WriteLine("Workload,ReferenceString,Frames,Policy,ClockStart,ResetInterval,ModifiedPages,Hit,Fault,Migration,FaultRate,EstimatedDelayMs");

            for (int w = 0; w < refs.Length; w++)
            {
                foreach (int frameSize in frames)
                {
                    foreach (Core.ReplacementPolicy policy in policies)
                    {
                        Run(names[w], refs[w], frameSize, policy, 1, 4, "AD");
                    }
                }
            }

            foreach (int clockStart in new[] { 1, 2, 3, 4 })
            {
                Run("ClockParam", "ABCDEABCDA", 4, Core.ReplacementPolicy.SecondChance, clockStart, 4, "AD");
            }

            foreach (int resetInterval in new[] { 2, 3, 4, 6 })
            {
                Run("ResetParam", "ABCDABEFABGHABCD", 4, Core.ReplacementPolicy.NUR_01_First, 1, resetInterval, "AD");
                Run("ResetParam", "ABCDABEFABGHABCD", 4, Core.ReplacementPolicy.NUR_10_First, 1, resetInterval, "AD");
            }

            foreach (string modifiedPages in new[] { "A", "AD", "BDF", "" })
            {
                Run("ModifiedParam", "ABCDABEFABGHABCD", 4, Core.ReplacementPolicy.NUR_01_First, 1, 4, modifiedPages);
                Run("ModifiedParam", "ABCDABEFABGHABCD", 4, Core.ReplacementPolicy.NUR_10_First, 1, 4, modifiedPages);
            }
        }
    }
}
