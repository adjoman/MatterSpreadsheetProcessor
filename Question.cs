using System;
using System.Collections.Generic;

namespace ImportMatterExcel
{
    public class Question
    {
        public string uuid { get; set; }
        public string type { get; set; }
        public string title { get; set; }
        public string section { get; set; }
        public List<object> tags { get; set; }
        public string question { get; set; }
        public List<Answer> answers { get; set; }
        public bool? required { get; set; }
        public string regex { get; set; }
    }
}
