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
            Optimal,
            LRU,
            SecondChance,
            LRFULite
        }

        public const int PageFaultDelayUnitMs = 10;
        private const double LrfuDecayFactor = 0.85;

        private int clockHand;
        private int arrivalSequence;
        private readonly ReplacementPolicy policy;
        private readonly Dictionary<char, int> lastUsed;
        private readonly Dictionary<char, int> frequency;
        private readonly Dictionary<char, bool> referenceBits;
        private readonly Dictionary<char, int> arrivalOrder;
        private readonly Dictionary<char, double> lrfuScores;

        public int p_frame_size;
        public List<Page> frame_window;
        public List<Page> pageHistory;

        public int hit;
        public int fault;
        public int migration;

        public Core(int get_frame_size, ReplacementPolicy get_policy)
        {
            if (get_frame_size <= 0)
            {
                throw new ArgumentException("Frame size must be greater than zero.");
            }

            Page.CREATE_ID = 0;
            this.clockHand = 0;
            this.arrivalSequence = 0;
            this.policy = get_policy;
            this.p_frame_size = get_frame_size;
            this.frame_window = new List<Page>();
            this.pageHistory = new List<Page>();
            this.lastUsed = new Dictionary<char, int>();
            this.frequency = new Dictionary<char, int>();
            this.referenceBits = new Dictionary<char, bool>();
            this.arrivalOrder = new Dictionary<char, int>();
            this.lrfuScores = new Dictionary<char, double>();
        }

        public Page.STATUS Operate(char data, int index, string referenceString)
        {
            Page newPage = CreateHistoryPage(data);

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
                UpdateMetadataOnHit(data, index);
                newPage.detail = "Hit: page already exists in frame " + newPage.loc;
            }
            else
            {
                this.fault++;
                UpdateMetadataOnFault(data, index);

                if (this.frame_window.Count >= this.p_frame_size)
                {
                    int victimIndex = SelectVictim(index, referenceString);
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
                    AddMetadata(data, index);
                    this.migration++;

                    if (this.policy == ReplacementPolicy.SecondChance)
                    {
                        this.clockHand = (victimIndex + 1) % this.p_frame_size;
                    }
                }
                else
                {
                    newPage.status = Page.STATUS.PAGEFAULT;
                    newPage.loc = this.frame_window.Count + 1;
                    newPage.detail = "Fault: inserted into empty frame " + newPage.loc;
                    this.frame_window.Add(CreateResidentPage(data));
                    AddMetadata(data, index);
                }
            }

            newPage.frameSnapshot = BuildFrameSnapshot();
            this.pageHistory.Add(newPage);

            return newPage.status;
        }

        public static ReplacementPolicy ParsePolicy(string text)
        {
            switch ((text ?? "").Trim())
            {
                case "Optimal":
                    return ReplacementPolicy.Optimal;
                case "LRU":
                    return ReplacementPolicy.LRU;
                case "Second Chance":
                    return ReplacementPolicy.SecondChance;
                case "LRFU-Lite":
                    return ReplacementPolicy.LRFULite;
                case "FIFO":
                default:
                    return ReplacementPolicy.FIFO;
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

        private Page CreateHistoryPage(char data)
        {
            Page page = new Page();
            page.pid = Page.CREATE_ID++;
            page.data = data;
            page.hasVictim = false;
            page.victim = '\0';
            page.frameSnapshot = "";
            page.detail = "";
            return page;
        }

        private Page CreateResidentPage(char data)
        {
            Page page = new Page();
            page.data = data;
            return page;
        }

        private int SelectVictim(int currentIndex, string referenceString)
        {
            switch (this.policy)
            {
                case ReplacementPolicy.Optimal:
                    return SelectOptimalVictim(currentIndex, referenceString);
                case ReplacementPolicy.LRU:
                    return SelectLruVictim();
                case ReplacementPolicy.SecondChance:
                    return SelectSecondChanceVictim();
                case ReplacementPolicy.LRFULite:
                    return SelectLrfuLiteVictim();
                case ReplacementPolicy.FIFO:
                default:
                    return 0;
            }
        }

        private int SelectOptimalVictim(int currentIndex, string referenceString)
        {
            int victimIndex = 0;
            int farthestNextUse = -1;

            for (int i = 0; i < this.frame_window.Count; i++)
            {
                char frameData = this.frame_window[i].data;
                int nextUse = int.MaxValue;

                for (int j = currentIndex + 1; j < referenceString.Length; j++)
                {
                    if (referenceString[j] == frameData)
                    {
                        nextUse = j;
                        break;
                    }
                }

                if (nextUse > farthestNextUse)
                {
                    farthestNextUse = nextUse;
                    victimIndex = i;
                }
            }

            return victimIndex;
        }

        private int SelectLruVictim()
        {
            int victimIndex = 0;
            int oldestUse = int.MaxValue;

            for (int i = 0; i < this.frame_window.Count; i++)
            {
                char frameData = this.frame_window[i].data;
                int usedAt = this.lastUsed.ContainsKey(frameData) ? this.lastUsed[frameData] : -1;

                if (usedAt < oldestUse)
                {
                    oldestUse = usedAt;
                    victimIndex = i;
                }
            }

            return victimIndex;
        }

        private int SelectSecondChanceVictim()
        {
            while (true)
            {
                char frameData = this.frame_window[this.clockHand].data;
                bool referenced = this.referenceBits.ContainsKey(frameData) && this.referenceBits[frameData];

                if (!referenced)
                {
                    return this.clockHand;
                }

                this.referenceBits[frameData] = false;
                this.clockHand = (this.clockHand + 1) % this.p_frame_size;
            }
        }

        private int SelectLrfuLiteVictim()
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

            return victimIndex;
        }

        private void UpdateMetadataOnHit(char data, int index)
        {
            this.lastUsed[data] = index;

            if (this.frequency.ContainsKey(data))
            {
                this.frequency[data]++;
            }

            if (this.referenceBits.ContainsKey(data))
            {
                this.referenceBits[data] = true;
            }

            if (this.lrfuScores.ContainsKey(data))
            {
                this.lrfuScores[data] += 1.0;
            }
        }

        private void UpdateMetadataOnFault(char data, int index)
        {
            this.lastUsed[data] = index;
        }

        private void AddMetadata(char data, int index)
        {
            this.lastUsed[data] = index;
            this.frequency[data] = 1;
            this.referenceBits[data] = false;
            this.arrivalOrder[data] = this.arrivalSequence++;
            this.lrfuScores[data] = 1.0;
        }

        private void RemoveMetadata(char data)
        {
            this.lastUsed.Remove(data);
            this.frequency.Remove(data);
            this.referenceBits.Remove(data);
            this.arrivalOrder.Remove(data);
            this.lrfuScores.Remove(data);
        }

        private void DecayLrfuScores()
        {
            List<char> keys = this.lrfuScores.Keys.ToList();

            foreach (char key in keys)
            {
                this.lrfuScores[key] *= LrfuDecayFactor;
            }
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
    }
}
