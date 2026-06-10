using System;
using System.Collections.Generic;
using System.Linq;

namespace Memory_Policy_Simulator
{
    class Core
    {
        public enum ReplacementPolicy
        {
            FIFO,
            NUR_01_First,
            NUR_10_First,
            SecondChance,
            LRFULite
        }

        public const int PageFaultDelayUnitMs = 10;
        private const double LrfuDecayFactor = 0.85;

        private int clockHand;
        private readonly int referenceResetInterval;
        private readonly ReplacementPolicy policy;
        private readonly HashSet<char> modifiedInputPages;
        private readonly Dictionary<char, bool> referenceBits;
        private readonly Dictionary<char, bool> modifiedBits;
        private readonly Dictionary<char, double> lrfuScores;

        public int p_frame_size;
        public List<Page> frame_window;
        public List<Page> pageHistory;

        public int hit;
        public int fault;
        public int migration;

        public Core(int get_frame_size, ReplacementPolicy get_policy)
            : this(get_frame_size, get_policy, 0, 4, "")
        {
        }

        public Core(int get_frame_size, ReplacementPolicy get_policy, int clock_start, int reset_interval, string modified_pages)
        {
            if (get_frame_size <= 0)
            {
                throw new ArgumentException("Frame size must be greater than zero.");
            }

            Page.CREATE_ID = 0;
            this.p_frame_size = get_frame_size;
            this.policy = get_policy;
            this.clockHand = NormalizeClockStart(clock_start, get_frame_size);
            this.referenceResetInterval = reset_interval <= 0 ? 0 : reset_interval;
            this.modifiedInputPages = new HashSet<char>((modified_pages ?? "").Where(x => !Char.IsWhiteSpace(x)));
            this.frame_window = new List<Page>();
            this.pageHistory = new List<Page>();
            this.referenceBits = new Dictionary<char, bool>();
            this.modifiedBits = new Dictionary<char, bool>();
            this.lrfuScores = new Dictionary<char, double>();
        }

        public Page.STATUS Operate(char data, int index, string referenceString)
        {
            Page newPage = CreateHistoryPage(data);

            ApplyPeriodicReferenceReset(index, newPage);

            if (this.policy == ReplacementPolicy.LRFULite)
            {
                DecayLrfuScores();
            }

            int hitIndex = this.frame_window.FindIndex(x => x.data == data);

            if (hitIndex >= 0)
            {
                newPage.status = Page.STATUS.HIT;
                newPage.loc = hitIndex + 1;
                this.hit++;
                MarkReferenced(data);
                newPage.detail = "Hit: page already exists in frame " + newPage.loc;
            }
            else
            {
                this.fault++;

                if (this.frame_window.Count >= this.p_frame_size)
                {
                    int victimIndex = SelectVictim(newPage);
                    Page victim = this.frame_window[victimIndex];

                    newPage.status = Page.STATUS.MIGRATION;
                    newPage.loc = victimIndex + 1;
                    newPage.hasVictim = true;
                    newPage.victim = victim.data;
                    newPage.detail = "Fault: replaced " + victim.data + " in frame " + newPage.loc;

                    RemoveMetadata(victim.data);

                    if (this.policy == ReplacementPolicy.FIFO)
                    {
                        this.frame_window.RemoveAt(victimIndex);
                        this.frame_window.Add(CreateResidentPage(data));
                        newPage.loc = this.p_frame_size;
                    }
                    else
                    {
                        this.frame_window[victimIndex] = CreateResidentPage(data);
                    }

                    AddMetadata(data);
                    this.migration++;
                }
                else
                {
                    newPage.status = Page.STATUS.PAGEFAULT;
                    newPage.loc = this.frame_window.Count + 1;
                    newPage.detail = "Fault: inserted into empty frame " + newPage.loc;
                    this.frame_window.Add(CreateResidentPage(data));
                    AddMetadata(data);
                }
            }

            newPage.frameSnapshot = BuildFrameSnapshot();
            newPage.algorithmState = AppendState(newPage.algorithmState, BuildAlgorithmState());
            this.pageHistory.Add(newPage);

            return newPage.status;
        }

        public static ReplacementPolicy ParsePolicy(string text)
        {
            switch ((text ?? "").Trim())
            {
                case "NUR (0,1 first)":
                    return ReplacementPolicy.NUR_01_First;
                case "NUR (1,0 first)":
                    return ReplacementPolicy.NUR_10_First;
                case "Second Chance":
                    return ReplacementPolicy.SecondChance;
                case "LRFU-Lite":
                    return ReplacementPolicy.LRFULite;
                case "FIFO":
                default:
                    return ReplacementPolicy.FIFO;
            }
        }

        public string GetPolicyDescription()
        {
            switch (this.policy)
            {
                case ReplacementPolicy.NUR_01_First:
                    return "NUR order: (0,0) -> (0,1) -> (1,0) -> (1,1), R reset interval=" + this.referenceResetInterval;
                case ReplacementPolicy.NUR_10_First:
                    return "NUR order: (0,0) -> (1,0) -> (0,1) -> (1,1), R reset interval=" + this.referenceResetInterval;
                case ReplacementPolicy.SecondChance:
                    return "Second Chance clock start frame=" + (this.clockHand + 1);
                case ReplacementPolicy.LRFULite:
                    return "LRFU-Lite decay factor=" + LrfuDecayFactor;
                case ReplacementPolicy.FIFO:
                default:
                    return "FIFO keeps the original insertion order.";
            }
        }

        public List<Page> GetPageInfo(Page.STATUS status)
        {
            List<Page> pages = new List<Page>();

            foreach (Page page in pageHistory)
            {
                if (page.status == status)
                {
                    pages.Add(page);
                }
            }

            return pages;
        }

        public int GetEstimatedPageFaultDelay()
        {
            return this.fault * PageFaultDelayUnitMs;
        }

        private static int NormalizeClockStart(int clockStart, int frameSize)
        {
            if (clockStart <= 0)
            {
                return 0;
            }

            return (clockStart - 1) % frameSize;
        }

        private Page CreateHistoryPage(char data)
        {
            Page page = new Page();
            page.pid = Page.CREATE_ID++;
            page.data = data;
            page.hasVictim = false;
            page.victim = '\0';
            page.frameSnapshot = "";
            page.detail = "";
            page.algorithmState = "";
            return page;
        }

        private Page CreateResidentPage(char data)
        {
            Page page = new Page();
            page.data = data;
            return page;
        }

        private int SelectVictim(Page historyPage)
        {
            switch (this.policy)
            {
                case ReplacementPolicy.NUR_01_First:
                    return SelectNurVictim(false, historyPage);
                case ReplacementPolicy.NUR_10_First:
                    return SelectNurVictim(true, historyPage);
                case ReplacementPolicy.SecondChance:
                    return SelectSecondChanceVictim(historyPage);
                case ReplacementPolicy.LRFULite:
                    return SelectLrfuLiteVictim(historyPage);
                case ReplacementPolicy.FIFO:
                default:
                    historyPage.algorithmState = AppendState(historyPage.algorithmState, "FIFO victim: oldest page");
                    return 0;
            }
        }

        private int SelectNurVictim(bool preferReferencedClean, Page historyPage)
        {
            int[][] order = preferReferencedClean
                ? new int[][] { new[] { 0, 0 }, new[] { 1, 0 }, new[] { 0, 1 }, new[] { 1, 1 } }
                : new int[][] { new[] { 0, 0 }, new[] { 0, 1 }, new[] { 1, 0 }, new[] { 1, 1 } };

            foreach (int[] targetClass in order)
            {
                for (int i = 0; i < this.frame_window.Count; i++)
                {
                    char frameData = this.frame_window[i].data;
                    int r = GetReferenceBit(frameData) ? 1 : 0;
                    int m = GetModifiedBit(frameData) ? 1 : 0;

                    if (r == targetClass[0] && m == targetClass[1])
                    {
                        historyPage.algorithmState = AppendState(
                            historyPage.algorithmState,
                            "NUR victim class=(" + r + "," + m + "), priority=" + BuildNurOrderText(order));
                        return i;
                    }
                }
            }

            return 0;
        }

        private int SelectSecondChanceVictim(Page historyPage)
        {
            List<string> scans = new List<string>();

            while (true)
            {
                char frameData = this.frame_window[this.clockHand].data;
                bool referenced = GetReferenceBit(frameData);
                scans.Add("F" + (this.clockHand + 1) + ":" + frameData + "/R=" + (referenced ? "1" : "0"));

                if (!referenced)
                {
                    int selected = this.clockHand;
                    this.clockHand = (this.clockHand + 1) % this.p_frame_size;
                    historyPage.algorithmState = AppendState(
                        historyPage.algorithmState,
                        "Second Chance scan=" + string.Join(" -> ", scans.ToArray()) + ", nextClock=F" + (this.clockHand + 1));
                    return selected;
                }

                this.referenceBits[frameData] = false;
                this.clockHand = (this.clockHand + 1) % this.p_frame_size;
            }
        }

        private int SelectLrfuLiteVictim(Page historyPage)
        {
            int victimIndex = 0;
            double lowestScore = double.MaxValue;

            for (int i = 0; i < this.frame_window.Count; i++)
            {
                char frameData = this.frame_window[i].data;
                double score = this.lrfuScores.ContainsKey(frameData) ? this.lrfuScores[frameData] : 0.0;

                if (score < lowestScore)
                {
                    lowestScore = score;
                    victimIndex = i;
                }
            }

            historyPage.algorithmState = AppendState(historyPage.algorithmState, "LRFU victim score=" + lowestScore.ToString("0.###"));
            return victimIndex;
        }

        private void MarkReferenced(char data)
        {
            this.referenceBits[data] = true;

            if (this.modifiedInputPages.Contains(data))
            {
                this.modifiedBits[data] = true;
            }

            if (this.lrfuScores.ContainsKey(data))
            {
                this.lrfuScores[data] += 1.0;
            }
        }

        private void AddMetadata(char data)
        {
            this.referenceBits[data] = true;
            this.modifiedBits[data] = this.modifiedInputPages.Contains(data);
            this.lrfuScores[data] = 1.0;
        }

        private void RemoveMetadata(char data)
        {
            this.referenceBits.Remove(data);
            this.modifiedBits.Remove(data);
            this.lrfuScores.Remove(data);
        }

        private void ApplyPeriodicReferenceReset(int index, Page historyPage)
        {
            if (this.referenceResetInterval <= 0 || index <= 0 || index % this.referenceResetInterval != 0)
            {
                return;
            }

            List<char> keys = this.referenceBits.Keys.ToList();

            foreach (char key in keys)
            {
                this.referenceBits[key] = false;
            }

            historyPage.algorithmState = AppendState(historyPage.algorithmState, "R bits reset before step " + (index + 1));
        }

        private void DecayLrfuScores()
        {
            List<char> keys = this.lrfuScores.Keys.ToList();

            foreach (char key in keys)
            {
                this.lrfuScores[key] *= LrfuDecayFactor;
            }
        }

        private bool GetReferenceBit(char data)
        {
            return this.referenceBits.ContainsKey(data) && this.referenceBits[data];
        }

        private bool GetModifiedBit(char data)
        {
            return this.modifiedBits.ContainsKey(data) && this.modifiedBits[data];
        }

        private string BuildFrameSnapshot()
        {
            List<string> frames = new List<string>();

            for (int i = 0; i < this.p_frame_size; i++)
            {
                frames.Add(i < this.frame_window.Count ? this.frame_window[i].data.ToString() : "-");
            }

            return string.Join(" ", frames.ToArray());
        }

        private string BuildAlgorithmState()
        {
            switch (this.policy)
            {
                case ReplacementPolicy.NUR_01_First:
                case ReplacementPolicy.NUR_10_First:
                    return BuildBitSnapshot();
                case ReplacementPolicy.SecondChance:
                    return "clock=F" + (this.clockHand + 1) + "; " + BuildBitSnapshot();
                case ReplacementPolicy.LRFULite:
                    return BuildScoreSnapshot();
                case ReplacementPolicy.FIFO:
                default:
                    return "fifo-order=[" + BuildFrameSnapshot() + "]";
            }
        }

        private string BuildBitSnapshot()
        {
            List<string> states = new List<string>();

            foreach (Page page in this.frame_window)
            {
                states.Add(page.data + "(R=" + (GetReferenceBit(page.data) ? "1" : "0") +
                    ",M=" + (GetModifiedBit(page.data) ? "1" : "0") + ")");
            }

            return string.Join(" ", states.ToArray());
        }

        private string BuildScoreSnapshot()
        {
            List<string> states = new List<string>();

            foreach (Page page in this.frame_window)
            {
                double score = this.lrfuScores.ContainsKey(page.data) ? this.lrfuScores[page.data] : 0.0;
                states.Add(page.data + "=" + score.ToString("0.###"));
            }

            return "scores{" + string.Join(", ", states.ToArray()) + "}";
        }

        private static string BuildNurOrderText(int[][] order)
        {
            return string.Join(">", order.Select(x => "(" + x[0] + "," + x[1] + ")").ToArray());
        }

        private static string AppendState(string current, string message)
        {
            if (String.IsNullOrEmpty(current))
            {
                return message;
            }

            return current + "; " + message;
        }
    }
}
