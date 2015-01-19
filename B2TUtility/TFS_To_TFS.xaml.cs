using System;
using System.Text;
using System.Windows;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Client;
using System.Configuration;

namespace xmlRpcExample
{
    /// <summary>
    /// Interaction logic for TFS_To_TFS.xaml
    /// </summary>
    public partial class TFS_To_TFS : Window
    {
        private TfsTeamProjectCollection from_tfsCol;
        WorkItemStore _from_store;
        string _bugState;
        string _bugResolution;

        private TfsTeamProjectCollection to_tfsCol;
        WorkItemStore _to_store;
        int _newBugId;

        public TFS_To_TFS()
        {
            InitializeComponent();

            //Connect to source TFS
            var tfsUriSetting = ConfigurationManager.AppSettings.Get("From_TFSUri");
            var tfsUri = new Uri(tfsUriSetting);
            from_tfsCol = new TfsTeamProjectCollection(tfsUri);
            _from_store = (WorkItemStore)from_tfsCol.GetService(typeof(WorkItemStore));

            //Connect to target TFS
            tfsUriSetting = ConfigurationManager.AppSettings.Get("To_TFSUri");
            tfsUri = new Uri(tfsUriSetting);
            to_tfsCol = new TfsTeamProjectCollection(tfsUri);
            _to_store = (WorkItemStore)to_tfsCol.GetService(typeof(WorkItemStore));
          
        }

        private void btnTransfer_Click(object sender, RoutedEventArgs e)
        {
            int bugId = int.Parse(txtBugId.Text.Trim());
            string from_projectName = ConfigurationManager.AppSettings["From_ProjectName"];
            string from_workItemTypeName = ConfigurationManager.AppSettings["From_WorkItemTypeName"];
            
            BuzillaBug _fromBug= GetWorkItem(_from_store, from_projectName, from_workItemTypeName, bugId);


            string to_projectName = ConfigurationManager.AppSettings["To_ProjectName"];
            string to_workItemTypeName = ConfigurationManager.AppSettings["To_WorkItemTypeName"];
            
            WriteWorkItem(_fromBug, _to_store, to_projectName, to_workItemTypeName);

            MessageBox.Show("DONE");
        }

        private BuzillaBug GetWorkItem(WorkItemStore from_store, string from_projectName, string from_workItemTypeName, int bugId)
        {

            BuzillaBug bug = new BuzillaBug();
            var tfsBugInfo = from_store.GetWorkItem(bugId);
            // Open it up for editing.  (Sometimes PartialOpen() works too and takes less time.)

            bug.Title = tfsBugInfo.Title;

            bug.BugStatus = tfsBugInfo.Fields["System.State"].Value.ToString();

            bug.Resolution = tfsBugInfo.Fields["Bugzilla.Resolution"].Value.ToString();

            bug.AssignedTo = tfsBugInfo.Fields["System.AssignedTo"].Value.ToString();

            bug.DocImpact = tfsBugInfo.Fields["Bugzilla.DocImpact"].Value.ToString();

            bug.EstimatedTime = tfsBugInfo.Fields["Bugzilla.EstimatedTime"].Value.ToString();

            bug.RemainingTime = tfsBugInfo.Fields["Bugzilla.RemainingTime"].Value.ToString();

            bug.ActualTime = tfsBugInfo.Fields["Bugzilla.ActualTime"].Value.ToString();

            bug.WhiteBoard = tfsBugInfo.Fields["Bugzilla.Whiteboard"].Value.ToString();

            bug.ExternalReporter = tfsBugInfo.Fields["Bugzilla.ExternalReporter"].Value.ToString();

            bug.CreatedBy = tfsBugInfo.Fields["System.CreatedBy"].Value.ToString();

            bug.ExternalBugId = tfsBugInfo.Fields["Bugzilla.ExternalBugId"].Value.ToString();

            bug.Severity = tfsBugInfo.Fields["Bugzilla.BugSeverity"].Value.ToString();

            bug.Priority = tfsBugInfo.Fields["Bugzilla.BugPriority"].Value.ToString();

            bug.Product = tfsBugInfo.Fields["Bugzilla.Product"].Value.ToString();

            bug.Component = tfsBugInfo.Fields["Bugzilla.Component"].Value.ToString();

            bug.Version = tfsBugInfo.Fields["Bugzilla.Version"].Value.ToString();

            bug.TargetMilestone = tfsBugInfo.Fields["Bugzilla.TargetMilestone"].Value.ToString();

            bug.CreatedDateTime = DateTime.Parse( tfsBugInfo.Fields["System.CreatedDate"].Value.ToString());
            
            bug.StepsToReproduce = tfsBugInfo.Fields["Microsoft.VSTS.TCM.ReproSteps"].Value.ToString();

            bug.ReportedIn = tfsBugInfo.Fields["Bugzilla.ReportedIn"].Value.ToString();

            bug.FixedIn = tfsBugInfo.Fields["Bugzilla.FixedIn"].Value.ToString();

            bug.Platform = tfsBugInfo.Fields["Bugzilla.Platform"].Value.ToString();

            bug.OperatingSystem = tfsBugInfo.Fields["Bugzilla.OS"].Value.ToString();

           //bug.Description = GetWorkItemHistory(tfsBugInfo);

            foreach (Attachment atc in tfsBugInfo.Attachments)
            {
                System.Net.WebClient request = new System.Net.WebClient(); ;

                request.Credentials = System.Net.CredentialCache.DefaultCredentials;

                request.DownloadFile(atc.Uri, System.IO.Path.Combine(@"C:\Attachments", atc.Name));

                var attachment = new Attachment(System.IO.Path.Combine(@"C:\Attachments", atc.Name), atc.Comment);

                bug.Attachments.Add(attachment);
            }
                                    
            bug.CcList = tfsBugInfo.Fields["Bugzilla.CCList2"].Value.ToString();
    
            return bug;
           
        }

        private string GetWorkItemHistory(WorkItem tfsBugInfo)
        {
            StringBuilder historyBuilder = new StringBuilder();
            foreach (Revision revision in tfsBugInfo.Revisions)
            {
                //String history = (String)revision.Fields["History"].Value;
                //historyBuilder.Append(history);

                foreach (Field item in revision.Fields)
                {
                    historyBuilder.Append(item.Name + " : " + item.Value + " ");
                }
                
            }

            return historyBuilder.ToString();

        }

        private bool WriteWorkItem(BuzillaBug bug, WorkItemStore store, string projectName, string workItemTypeName)
        {
            bool wroteToTfs = false;


            Project project;

            for (int i = 0; i < store.Projects.Count; i++)
            {
                if (store.Projects[i].Name == projectName)
                {
                    project = store.Projects[i];

                    for (int j = 0; j < project.WorkItemTypes.Count; j++)
                    {
                        if (project.WorkItemTypes[j].Name == workItemTypeName)
                        {
                            WorkItemType workItemType = project.WorkItemTypes[j];
                            
                            SaveBug(bug, workItemType);                            
                            
                            wroteToTfs = true;
                            break;
                        }
                    }
                    break;
                }
            }


            if (!wroteToTfs)
            {
                MessageBox.Show("Project or WorkItem not found. Please make sure the App.Config is correct.");
            }
            return wroteToTfs;

        }

        private void SaveBug(BuzillaBug bug, WorkItemType workItemType)
        {
            var newWI = workItemType.NewWorkItem();
                        
            newWI.Title = bug.Title;

            //newWI.Fields["Bugzilla.ID"].Value = bug.bugzilla_BugId;

            //**** newWI.Fields["Bugzilla.Status"].Value = bug.BugStatus;

            //**** newWI.Fields["Bugzilla.Resolution"].Value = bug.Resolution;

            newWI.Fields["System.AssignedTo"].Value = bug.AssignedTo;

            if (bug.DocImpact != null)
                newWI.Fields["Bugzilla.DocImpact"].Value = bug.DocImpact;

            newWI.Fields["Bugzilla.Whiteboard"].Value = bug.WhiteBoard;

            if (bug.ExternalReporter != null)
                newWI.Fields["Bugzilla.ExternalReporter"].Value = bug.ExternalReporter;

            if (bug.ExternalBugId != null)
                newWI.Fields["Bugzilla.ExternalBugId"].Value = bug.ExternalBugId;

            newWI.Fields["Bugzilla.BugSeverity"].Value = bug.Severity;

            newWI.Fields["Bugzilla.BugPriority"].Value = bug.Priority;

            newWI.Fields["Bugzilla.Product"].Value = bug.Product;

            newWI.Fields["Bugzilla.Component"].Value = bug.Component;

            newWI.Fields["Bugzilla.Version"].Value = bug.Version;

            newWI.Fields["Bugzilla.TargetMilestone"].Value = bug.TargetMilestone;

            var stepsToRepro = "[Imported on behalf of " + bug.CreatedBy +" ]" + Environment.NewLine + bug.StepsToReproduce; 
            newWI.Fields["Microsoft.VSTS.TCM.ReproSteps"].Value = stepsToRepro;

            if (bug.ReportedIn != null)
                newWI.Fields["Bugzilla.ReportedIn"].Value = bug.ReportedIn;

            if (bug.FixedIn != null)
                newWI.Fields["Bugzilla.FixedIn"].Value = bug.FixedIn;

            newWI.Fields["Bugzilla.Platform"].Value = bug.Platform;

            newWI.Fields["Bugzilla.OS"].Value = bug.OperatingSystem;


            //newWI.History = bug.Description;

            if (bug.Attachments != null && bug.Attachments.Count > 0)
            {
                foreach (Attachment attachment in bug.Attachments)
                {
                    newWI.Attachments.Add(attachment);
                }
            }

            newWI.Fields["Bugzilla.CCList2"].Value = bug.CcList;

            newWI.Save();
            
            _newBugId = newWI.Id;

            if (bug.BugStatus.ToLower() != "new")
            {
                _bugState = bug.BugStatus;
                _bugResolution = bug.Resolution;
                newWI.Close();
                UpdateState();

            }
        }

        private void UpdateState()
        {
            var to_store = (WorkItemStore)to_tfsCol.GetService(typeof(WorkItemStore));
            //Note: The Project and WorkItemType has to be specifically selected in final implementation
            string to_projectName = ConfigurationManager.AppSettings["To_ProjectName"];
            string to_workItemTypeName = ConfigurationManager.AppSettings["To_WorkItemTypeName"];

            var workItem = to_store.GetWorkItem(_newBugId);
            workItem.Open();

            workItem.Fields["System.State"].Value = _bugState;
            workItem.Fields["Bugzilla.Resolution"].Value = _bugResolution;
         
            workItem.Save();
            workItem.Close();

        }

    }
}
