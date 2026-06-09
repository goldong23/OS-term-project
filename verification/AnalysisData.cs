using System;

namespace Memory_Policy_Simulator
{
    class AnalysisData
    {
        static void Run(string workloadName, string referenceString, int frameSize, Core.ReplacementPolicy policy)
        {
            Core core = new Core(frameSize, policy);

            for (int i = 0; i < referenceString.Length; i++)
            {
                core.Operate(referenceString[i], i, referenceString);
            }

            double faultRate = Math.Round((double)core.fault / referenceString.Length * 100.0, 2);
            Console.WriteLine(workloadName + "," + referenceString + "," + frameSize + "," + policy + "," + core.hit + "," + core.fault + "," + core.migration + "," + faultRate + "," + core.GetEstimatedPageFaultDelay());
        }

        static void Main()
        {
            string[] names = { "Textbook", "Locality", "SequentialScan", "Mixed", "FrequencyBias" };
            string[] refs = { "123412512345", "ABCABCABCDEFDEFABC", "ABCDEFGHIJKL", "ABCDABEFABGHABCD", "AAAABCAAADEFAAA" };
            int[] frames = { 3, 4, 5 };
            Core.ReplacementPolicy[] policies =
            {
                Core.ReplacementPolicy.FIFO,
                Core.ReplacementPolicy.Optimal,
                Core.ReplacementPolicy.LRU,
                Core.ReplacementPolicy.SecondChance,
                Core.ReplacementPolicy.LRFULite
            };

            Console.WriteLine("Workload,ReferenceString,Frames,Policy,Hit,Fault,Migration,FaultRate,EstimatedDelayMs");

            for (int w = 0; w < refs.Length; w++)
            {
                foreach (int frameSize in frames)
                {
                    foreach (Core.ReplacementPolicy policy in policies)
                    {
                        Run(names[w], refs[w], frameSize, policy);
                    }
                }
            }
        }
    }
}
