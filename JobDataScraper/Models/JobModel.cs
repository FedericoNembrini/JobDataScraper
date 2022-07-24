using System.Collections.Generic;

namespace JobDataScraper.Models
{
    public class JobModel
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }

        public string Url { get; set; }

        public List<JobTask> JobTasks { get; set; } = new List<JobTask>();

        public List<JobKnowledge> JobKnowledges { get; set; } = new List<JobKnowledge>();

        public List<JobSkill> JobSkills { get; set; } = new List<JobSkill>();

        public List<JobExample> JobExamples { get; set; } = new List<JobExample>();

        public List<JobAttitude> JobAttitudes { get; set; } = new List<JobAttitude>();

        public List<JobGeneralisedActivity> JobGeneralisedActivities { get; set; } = new List<JobGeneralisedActivity>();

        public List<JobStyle> JobStyles { get; set; } = new List<JobStyle>();

        public List<JobEQF> JobEQFs { get; set; } = new List<JobEQF>();
    }

    public class JobTask
    {
        public string Name { get; set; }

        public decimal Importance { get; set; }

        public decimal Frequency { get; set; }
    }

    public class JobKnowledge
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public decimal Importance { get; set; }

        public decimal Complexity { get; set; }
    }

    public class JobSkill
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public decimal Importance { get; set; }

        public decimal Complexity { get; set; }
    }

    public class JobExample
    {
        public string Name { get; set; }
    }

    public class JobAttitude
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public decimal Importance { get; set; }

        public decimal Complexity { get; set; }
    }

    public class JobGeneralisedActivity
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public decimal Importance { get; set; }

        public decimal Complexity { get; set; }
    }

    public class JobStyle
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public decimal Importance { get; set; }
    }

    public class JobEQF
    {
        public string Name { get; set; }

        public decimal Value { get; set; }
    }
}
