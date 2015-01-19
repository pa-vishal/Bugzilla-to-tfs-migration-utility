using System;
using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace xmlRpcExample
{
    public class BuzillaBug
    {

        public BuzillaBug()
        {
            Attachments = new List<Attachment>();
        }

        public int bugzilla_BugId { get; set; }

        public DateTime CreatedDateTime { get; set; }

        public string Title { get; set; }

        public DateTime ChangedDateTime { get; set; }

        public string Product { get; set; }

        public string Component { get; set; }

        public string Version { get; set; }

        public string Platform { get; set; }

        public string OperatingSystem { get; set; }

        public string BugStatus { get; set; }

        public string Resolution { get; set; }

        public string Priority { get; set; }

        public string Severity { get; set; }

        public string TargetMilestone { get; set; }

        //dependson
        public int DependsOn { get; set; }

        //blocked
        public int Blocks { get; set; }
        
        
        public string ReportedBy { get; set; }

        public string AssignedTo { get; set; }

        public string CcList { get; set; }

        public string EstimatedTime { get; set; }

        public string RemainingTime { get; set; }

        public string ActualTime { get; set; }

        public string ExternalBugId { get; set; }

        public string ExternalReporter { get; set; }

        public string ReportedIn { get; set; }

        public string FixedIn { get; set; }

        public string DocImpact { get; set; }

        public string Description { get; set; } //extend t include: comment_id? , who, bug_when, 

        public string DescriptionWho { get; set; }

        public string DescriptionWhen { get; set; }

        public List<Attachment> Attachments { get; set; } //extend to include :attachid, date, delta_ts, desc, attacher

        // new Attachment(@"C:\Documents and Settings\vispatil\Desktop\New Project\Bugzilla\ImportedBugs\3799\VJ22_aph.png");


        public string WhiteBoard { get; set; }

        public string CreatedBy { get; set; }

        public string StepsToReproduce { get; set; }
    }

}
