using System.Collections.Generic;

namespace JobDataScraper.Models
{
    public class AdaJobModel
    {
        public string Code { get; set; }

        public string Description { get; set; }

        public List<AdaTask> Tasks { get; set; } = new List<AdaTask>();

        public List<AdaJobAssociation> Associations { get; set; } = new List<AdaJobAssociation>();

        public string CategoryName { get; set; }

        public string CategoryDescription { get; set; }
    }

    public class AdaTask
    {
        public string Description { get; set; }

        public string ExpectedResultDescription { get; set; }
    }

    public class AdaJobAssociation
    {
        public string Code { get; set; }

        public string Name { get; set; }
    }
}
