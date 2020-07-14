using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceClaim.Models
{
    public class WeeklyGWPModel
    {
        [Description("Branch")]
        public string BranchName { get; set; }
        [Description("First Week Count")]
        public int FirstWeekCount { get; set; }
        [Description("First Week Value")]
        public decimal FirstWeekValue { get; set; }
        [Description("Second Week Count")]
        public int SecondWeekCount { get; set; }
        [Description("Second Week Value")]
        public decimal SecondWeekValue { get; set; }
        [Description("Third Week Count")]
        public int ThirdWeekCount { get; set; }
        [Description("Third Week Value")]
        public decimal ThirdWeekValue { get; set; }
        [Description("Fourth Week Count")]
        public int FourWeekCount { get; set; }
        [Description("Fourth Week Value")]
        public decimal FourWeekValue { get; set; }
        [Description("Total Month Count")]
        public int TotalMonthCount { get; set; }
        [Description("Total Month Value")]
        public decimal TotalMonthValue { get; set; }



    }
}
