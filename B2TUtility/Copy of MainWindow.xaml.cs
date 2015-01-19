using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml;
using BaseClassNameSpace.Web.BaseServices;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace xmlRpcExample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {

        #region Members

        private readonly BackgroundWorker _workerBugzillaToXml;
        private readonly BackgroundWorker _workerXmlToAttachment;
        private readonly BackgroundWorker _workerToTFS;

        private TfsTeamProjectCollection tfsCol;

        int _batch = 10;
        int _start = 1;
        private Cursor _oldCursor;

        #endregion

        #region Constructor
        public MainWindow()
        {
            InitializeComponent();
            //_server = new Server("rdsw-bugzilla.Department.myCompany.com", "bugzilla");
            //this.DataContext = this;

            // Background Worker for Bugzilla to Xml
            _workerBugzillaToXml = new BackgroundWorker { WorkerReportsProgress = true };
            _workerBugzillaToXml.DoWork += WorkerDoWork;
            _workerBugzillaToXml.RunWorkerCompleted += WorkerRunWorkerCompleted;
            _workerBugzillaToXml.ProgressChanged += WorkerProgressChanged;

            // Background Worker for Xml to attachment
            _workerXmlToAttachment = new BackgroundWorker();
            _workerXmlToAttachment.DoWork += _workerXmlToAttachment_DoWork;
            _workerXmlToAttachment.RunWorkerCompleted += _workerXmlToAttachment_RunWorkerCompleted;

            // Background Worker for xml&attachment to TFS
            _workerToTFS = new BackgroundWorker { WorkerReportsProgress = true };
            _workerToTFS.DoWork += _workerToTFS_DoWork;
            _workerToTFS.RunWorkerCompleted += _workerToTFS_RunWorkerCompleted;

            //Connect to TFS
            var tfsUriSetting = ConfigurationManager.AppSettings.Get("TFSUri");
            var tfsUri = new Uri(tfsUriSetting);
            tfsCol = new TfsTeamProjectCollection(tfsUri);

        }
        #endregion

        #region Bugzilla To Xml Methods

        void WorkerDoWork(object sender, DoWorkEventArgs e)
        {
            var httpBase = new HttpBaseClass("vishal.patil@Company.com", "vishal@bugzilla", "", 0,
                                                      "http://rdsw-bugzilla.Department.myCompany.com/bugzilla/");

            var bugzillaLoginNumberInCookie = ConfigurationManager.AppSettings.Get("BugzillaLoginNumberInCookie");
            var buzillaLoginCodeInCookie = ConfigurationManager.AppSettings.Get("BugzillaLoginCodeInCookie"); 
            string cookie = "Cookie: DEFAULTFORMAT=specific; LASTORDER=bug_status%2Cpriority%2Cassigned_to%2Cbug_id; Bugzilla_login=" + bugzillaLoginNumberInCookie + 
                                                                                                            "; Bugzilla_logincookie=" + buzillaLoginCodeInCookie;
            const string requestMethod = "GET";
            const string queryStringFirstPart = "http://rdsw-bugzilla.Department.myCompany.com/bugzilla/show_bug.cgi?id=";
            const string queryStringLastPart = "&ctype=xml";

            // Note, this is a rather ineffecient way to do things, but
            // currently Bugzilla WebService does not support a way to
            // retrieve what bug ids one can access. So, we start from "start", then
            // proceed until an error occurs (which can be non-existant bug,
            // or access error), or continue we reach "batch" bugs

            string outputFileLocation = ConfigurationManager.AppSettings.Get("OutputFileLocation");
            if (!outputFileLocation.EndsWith(@"\")) outputFileLocation += @"\";

            int bugIdToReportError = 0;
            string fileNameToReportError = string.Empty;

            while (_batch > 0 && _start <= _batch)
            {
                try
                {
                    foreach (var bugId in Seq(_start, _batch - _start + 1))
                    {
                        bugIdToReportError = bugId;

                        string finalResponse = httpBase.GetFinalResponse(queryStringFirstPart + bugId + queryStringLastPart,
                                                                         cookie, requestMethod, true);
                        string fileName = outputFileLocation + "ImportedBug_" + bugId + ".xml";
                        fileNameToReportError = fileName;
                        _workerBugzillaToXml.ReportProgress(bugId, "Writing File: " + fileName);

                        var xmlTextWriter = new XmlTextWriter(fileName, Encoding.UTF8);

                        Regex regex = new Regex(@">\s*<");
                        string cleanedXml = regex.Replace(finalResponse, "><");

                        const string dtdString = "<!DOCTYPE bugzilla SYSTEM \"http://rdsw-bugzilla.Department.myCompany.com/bugzilla/bugzilla.dtd\">";
                        cleanedXml = cleanedXml.Replace(dtdString, " ");
                        xmlTextWriter.WriteRaw(cleanedXml);
                        Thread.Sleep(500);
                        xmlTextWriter.Flush();
                        xmlTextWriter.Close();

                    }

                    _start += _batch;
                }
                catch (Exception)
                {
                    _workerBugzillaToXml.ReportProgress(bugIdToReportError, "Could'nt write file: " + fileNameToReportError);
                    // Reduce batch
                    // batch = batch / 2;
                    continue;
                }
            }
        }

        void WorkerProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;

            string fileName = e.UserState.ToString();
            if (fileName.StartsWith("Could'nt"))
                tbc.Foreground = Brushes.Red;

            tbc.Text = fileName;
        }

        void WorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.Value = progressBar.Maximum;
            Cursor = _oldCursor;
            btnWebRequest.IsEnabled = true;
            //MessageBox.Show("Bugs from Bugzilla are successfully written to Xml.");
            tbc.Text = "Bugs from Bugzilla are successfully written to Xml.";
        }

        private void Web_Request_Click(object sender, RoutedEventArgs e)
        {
            string outputFileLocation = ConfigurationManager.AppSettings.Get("OutputFileLocation");
            if (!Directory.Exists(outputFileLocation))
                Directory.CreateDirectory(outputFileLocation);

            if (Directory.GetFiles(outputFileLocation).Length > 0)
                if (MessageBox.Show("There are existing files at :" + outputFileLocation + ", Do you want to delete them first?", "Clean up?", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    Directory.Delete(outputFileLocation, true);
                    Directory.CreateDirectory(outputFileLocation);
                }

            if (!int.TryParse(txtFrom.Text.Trim(), out _start))
                return;

            if (!int.TryParse(txtTo.Text.Trim(), out _batch))
                return;

            _oldCursor = Cursor;

            progressBar.Minimum = _start;
            progressBar.Maximum = _batch;
            progressBar.Value = 0;
            Cursor = Cursors.Wait;
            btnWebRequest.IsEnabled = false;
            _workerBugzillaToXml.RunWorkerAsync();
        }

        // Helper function - create a sequence of ints.
        static IEnumerable<int> Seq(int start, int count)
        {
            var res = new int[count];
            while (count > 0)
            {
                --count;
                res[count] = start + count;
            }
            return res;
        }
        #endregion

        #region Xml to Attachment
        void _workerXmlToAttachment_DoWork(object sender, DoWorkEventArgs e)
        {
            string bugFilePath = ConfigurationManager.AppSettings.Get("OutputFileLocation"); // @"C:\temp\ImportedBug_8204.xml";

            DirectoryInfo dirInfo = new DirectoryInfo(bugFilePath);

            foreach (var fileInfo in dirInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly))
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ProhibitDtd = false;
                XmlReader xmlReader = XmlReader.Create(fileInfo.FullName, settings);

                string bugId = string.Empty;
                string fileName = string.Empty;
                string attachId = string.Empty;
                try
                {
                    while (xmlReader.Read()) //ToDo: refactore to use XPath
                    {
                        if (xmlReader.NodeType == XmlNodeType.Element)
                        {
                            switch (xmlReader.Name)
                            {
                                case "bug_id":
                                    bugId = xmlReader.ReadString();
                                    break;
                                case "filename":
                                    fileName = xmlReader.ReadString();
                                    break;
                                case "attachid":
                                    attachId = xmlReader.ReadString();
                                    break;
                                case "data":
                                    string attachmentDataString = xmlReader.ReadString();

                                    if (attachmentDataString.Trim() != string.Empty)
                                    {
                                        string attachmentDir = fileInfo.DirectoryName + @"\" + bugId;
                                        if (!Directory.Exists(attachmentDir))
                                            Directory.CreateDirectory(attachmentDir);

                                        string pathForAttachment = attachmentDir + @"\" + fileName;

                                        CreateFileFromBytes(attachmentDataString, pathForAttachment, attachId);
                                    }

                                    break;
                            }
                        }
                    }
                }
                catch (XmlException)
                {
                    MessageBox.Show("Error in file : " + fileInfo.FullName + ". Skipped.");
                    continue;
                }
            }
        }

        private void CreateFileFromBytes(string attachmentDataString, string pathForAttachment, string attachId)
        {
            byte[] attachmentDataBytes = Convert.FromBase64String(attachmentDataString);
            if (File.Exists(pathForAttachment))
            {
                pathForAttachment += "_" + attachId;
            }
            FileStream stream = File.Create(pathForAttachment);
            stream.Write(attachmentDataBytes, 0, attachmentDataBytes.Length);
            stream.Close();
        }

        void _workerXmlToAttachment_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.IsIndeterminate = false;
            Cursor = _oldCursor;
            // MessageBox.Show("Attachments written successfully");
            tbc.Text = "Attachments written successfully.";
        }

        private void WriteAttachment_Click(object sender, RoutedEventArgs e)
        {
            _oldCursor = Cursor;
            Cursor = Cursors.Wait;
            progressBar.IsIndeterminate = true;
            _workerXmlToAttachment.RunWorkerAsync();
        }
        #endregion

        #region To TFS

        void _workerToTFS_DoWork(object sender, DoWorkEventArgs e)
        {

            var store = (WorkItemStore)tfsCol.GetService(typeof(WorkItemStore));

            //Note: The Project and WorkItemType has to be specifically selected in final implementation
            string projectName = ConfigurationManager.AppSettings["ProjectName"];
            string workItemTypeName = ConfigurationManager.AppSettings["WorkItemTypeName"];

            MessageBoxResult userDecision = MessageBox.Show("In App.Config, the ProjectName is : " + projectName
                                                         + " and \n\t the WorkItemTypeName is: " + workItemTypeName
                                                         + " Do you want to continue?", "Set App.Config", MessageBoxButton.YesNo
                                                         );

            if (userDecision == MessageBoxResult.No)
            {
                MessageBox.Show("After App.Config changes, please proceed only by clicking 'Write To TFS' button");
                return;             

            }

            string bugFilePath = ConfigurationManager.AppSettings.Get("OutputFileLocation"); // @"C:\temp\ImportedBug_8204.xml";

            DirectoryInfo dirInfo = new DirectoryInfo(bugFilePath);

            foreach (var fileInfo in dirInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly))
            {
                BuzillaBug bug;
                try
                {
                    bug = GetBug(dirInfo, fileInfo);
                }
                catch (Exception)
                {
                    //MessageBox.Show("Error reading the bug from xml. Please check bug xml: " + fileInfo.FullName + ". Utility will continue with next bug if any.",
                    //                "Failed write to TFS", MessageBoxButton.OK, MessageBoxImage.Error);

                    this.Dispatcher.Invoke(DispatcherPriority.Normal,
                                           new Action(
                                                        () => { tbc.Text += "\n Error reading file:" + fileInfo.FullName + " -> Bug import skipped." ; }
                                               )
                                           );
                    continue;
                }

                try
                {
                   
                    bool hasSucceded = WriteWorkItem(bug, store, projectName, workItemTypeName);
                    if (!hasSucceded)
                        break;
                    
                }
                catch (Exception)
                {
                    //MessageBox.Show("Error saving the bug to TFS. Please check connectivity to TFS Server or bug xml: " + fileInfo.FullName + ". Utility will continue with next bug if any.",
                    //                "Failed write to TFS", MessageBoxButton.OK, MessageBoxImage.Error);

                    this.Dispatcher.Invoke(DispatcherPriority.Normal,
                                            new Action(
                                                         () =>
                                                         {
                                                             //tbc.Foreground = Brushes.Red;
                                                                 tbc.Text += "\n Error importing bug from file:" + fileInfo.FullName + " -> Bug import skipped.";
                                                             }
                                                )
                                            );

                    Thread.Sleep(1000);
                    continue;
                }

            }
            e.Result = bugFilePath;
        }

        
         

        void _workerToTFS_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.IsIndeterminate = false;
            Cursor = _oldCursor;

            tbc.Text = "All bugs at location: " + e.Result + " are written to TFS successfully.";
        }

        private void WriteToTFS_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult userDecision = MessageBox.Show("Please make sure that the TFSWatcher service for email Alerts is 'Stopped'. Do you want to continue.", "Email Alerts!", MessageBoxButton.YesNo
                                                         );

            if (userDecision == MessageBoxResult.No)
            {
                MessageBox.Show("After stopping the TFSWatcher Service, please proceed only by clicking 'Write To TFS' button");
                return;
            }

            tfsCol.Connect(ConnectOptions.IncludeServices);
            Cursor = Cursors.Wait;

            progressBar.IsIndeterminate = true;
            _workerToTFS.RunWorkerAsync();
            tbc.Text = "Writting bugs to TFS.";

        }

        private BuzillaBug GetBug(DirectoryInfo dirInfo, FileInfo fileInfo)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(fileInfo.FullName);

            BuzillaBug bug = new BuzillaBug();

            //ToDo: will be nice if done using attributes on properties and reflection

            bug.bugzilla_BugId = int.Parse(doc.GetElementsByTagName("bug_id").Item(0).FirstChild.Value);

            bug.CreatedDateTime = DateTime.Parse(doc.GetElementsByTagName("creation_ts").Item(0).FirstChild.Value);

            bug.Title = doc.GetElementsByTagName("short_desc").Item(0).FirstChild.Value;

            bug.ChangedDateTime = DateTime.Parse(doc.GetElementsByTagName("delta_ts").Item(0).FirstChild.Value);

            if (doc.GetElementsByTagName("product").Count > 0 && doc.GetElementsByTagName("product").Item(0).FirstChild != null)
                bug.Product = doc.GetElementsByTagName("product").Item(0).FirstChild.Value;

            bug.Component = doc.GetElementsByTagName("component").Item(0).FirstChild.Value;

            bug.Version = doc.GetElementsByTagName("version").Item(0).FirstChild.Value;

            bug.Platform = doc.GetElementsByTagName("rep_platform").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("op_sys").Count > 0 && doc.GetElementsByTagName("op_sys").Item(0).FirstChild != null)
                bug.OperatingSystem = doc.GetElementsByTagName("op_sys").Item(0).FirstChild.Value;

            bug.BugStatus = doc.GetElementsByTagName("bug_status").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("resolution").Count > 0 && doc.GetElementsByTagName("resolution").Item(0).FirstChild != null)
                bug.Resolution = doc.GetElementsByTagName("resolution").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("priority").Count > 0 && doc.GetElementsByTagName("priority").Item(0).FirstChild != null)
                bug.Priority = doc.GetElementsByTagName("priority").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("bug_severity").Count > 0 && doc.GetElementsByTagName("bug_severity").Item(0).FirstChild != null)
                bug.Severity = doc.GetElementsByTagName("bug_severity").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("target_milestone").Count > 0 && doc.GetElementsByTagName("target_milestone").Item(0).FirstChild != null)
                bug.TargetMilestone = doc.GetElementsByTagName("target_milestone").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("dependson").Count > 0 && doc.GetElementsByTagName("dependson").Item(0).FirstChild != null)
                bug.DependsOn = int.Parse(doc.GetElementsByTagName("dependson").Item(0).FirstChild.Value);

            if (doc.GetElementsByTagName("blocked").Count > 0 && doc.GetElementsByTagName("blocked").Item(0).FirstChild != null)
                bug.Blocks = int.Parse(doc.GetElementsByTagName("blocked").Item(0).FirstChild.Value);

            bug.ReportedBy = doc.GetElementsByTagName("reporter").Item(0).FirstChild.Value;

            bug.AssignedTo = doc.GetElementsByTagName("assigned_to").Item(0).FirstChild.Value;

            //Export CCList
            bug.CcList = string.Empty;
            if (doc.GetElementsByTagName("cc").Count > 0)
                foreach (XmlElement node in doc.GetElementsByTagName("cc"))
                {
                    string emailId = node.InnerText;

                    if (emailId.Contains("@myCompany.com"))
                        emailId = emailId.Replace("@myCompany.com", "@Company.com");

                    bug.CcList += emailId + "\n ";
                }
            //if (bug.CcList != null) //Intentionally allowing it to be null (no init in ctor) 
            // bug.CcList = bug.CcList.Remove(bug.CcList.LastIndexOf(","));

            if (doc.GetElementsByTagName("estimated_time").Count > 0 && doc.GetElementsByTagName("estimated_time").Item(0).FirstChild != null)
                bug.EstimatedTime = doc.GetElementsByTagName("estimated_time").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("remaining_time").Count > 0 && doc.GetElementsByTagName("remaining_time").Item(0).FirstChild != null)
                bug.RemainingTime = doc.GetElementsByTagName("remaining_time").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("actual_time").Count > 0 && doc.GetElementsByTagName("actual_time").Item(0).FirstChild != null)
                bug.ActualTime = doc.GetElementsByTagName("actual_time").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("cf_extern_id").Count > 0 && doc.GetElementsByTagName("cf_extern_id").Item(0).FirstChild != null)
                bug.ExternalBugId = doc.GetElementsByTagName("cf_extern_id").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("cf_extern_reporter").Count > 0 && doc.GetElementsByTagName("cf_extern_reporter").Item(0).FirstChild != null)
                bug.ExternalReporter = doc.GetElementsByTagName("cf_extern_reporter").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("cf_release_badvers").Count > 0 && doc.GetElementsByTagName("cf_release_badvers").Item(0).FirstChild != null)
                bug.ReportedIn = doc.GetElementsByTagName("cf_release_badvers").Item(0).FirstChild.Value;

            if (doc.GetElementsByTagName("cf_release_fixed_in").Count > 0 && doc.GetElementsByTagName("cf_release_fixed_in").Item(0).FirstChild != null)
                bug.FixedIn = doc.GetElementsByTagName("cf_release_fixed_in").Item(0).FirstChild.Value;

            if (doc.SelectNodes(@"/bugzilla/bug/flag") != null)
            {
                XmlNodeList flag = doc.SelectNodes(@"/bugzilla/bug/flag");
                if (flag != null
                    && flag.Item(0) != null
                    && flag.Item(0).Attributes != null
                    && flag.Item(0).Attributes["name"].Value == "Doc_Impact?")
                {
                    bug.DocImpact = flag.Item(0).Attributes["status"].Value;
                }
            }

            //Description code goes here... [Description is being written as History in tfs..]
            XmlNodeList longDescNodes = doc.SelectNodes(@"/bugzilla/bug/long_desc");
            if (longDescNodes != null)
            {
                foreach (XmlLinkedNode xmlLinkedNode in longDescNodes)
                {
                    string fullDescription = GetFullDescription(bug, xmlLinkedNode);

                    bug.Description += fullDescription;
                }
            }

            //code for attachment goes here...
            var attachmentMetaData = new Dictionary<string, string>();

            if (doc.SelectNodes(@"/bugzilla/bug/attachment") != null)
            {
                XmlNodeList attachmentCommentsNodes = doc.SelectNodes(@"/bugzilla/bug/attachment");

                if (attachmentCommentsNodes != null)
                    foreach (XmlLinkedNode node in attachmentCommentsNodes)
                    {
                        string attchmentFileName;
                        string attachmentId;
                        string attachmentComment = GetAttachmentComment(node, out attchmentFileName, out attachmentId);
                        if (attchmentFileName != string.Empty)
                            if (!attachmentMetaData.ContainsKey(attchmentFileName))
                                attachmentMetaData.Add(attchmentFileName, attachmentComment);
                            else
                                attachmentMetaData.Add(attchmentFileName + "_" + attachmentId, attachmentComment);
                    }
            }
            var subDirs = dirInfo.GetDirectories();

            foreach (var subDir in subDirs)
            {
                if (subDir.Name == bug.bugzilla_BugId.ToString())
                {
                    foreach (var fileAttachments in subDir.GetFiles("*.*", SearchOption.TopDirectoryOnly))
                    {
                        var attachment = new Attachment(fileAttachments.FullName)
                        {
                            Comment = attachmentMetaData[fileAttachments.Name]
                        };
                        bug.Attachments.Add(attachment);
                    }

                }
            }
            return bug;
        }

        private string GetAttachmentComment(XmlLinkedNode node, out string attchmentFileName, out string attachmentId)
        {
            attchmentFileName = string.Empty;
            attachmentId = string.Empty;
            string attachmentComment = string.Empty;
            string attachmentDate = string.Empty;
            string attachmentChangeDate = string.Empty;
            foreach (XmlNode childNode in node.ChildNodes)
            {
                switch (childNode.Name)
                {
                    case "attachid":
                        attachmentId = childNode.InnerText;
                        break;
                    case "desc":
                        attachmentComment = childNode.InnerText;
                        break;
                    case "filename":
                        attchmentFileName = childNode.InnerText;
                        break;
                    case "date":
                        attachmentDate = childNode.InnerText;
                        break;
                    case "delta_ts":
                        attachmentChangeDate = childNode.InnerText;
                        break;

                }
            }
            attachmentComment = "Original Date: " + attachmentDate + ", Change Date: " +
                                attachmentChangeDate + " Comments:" + attachmentComment;
            return attachmentComment;
        }

        private string GetFullDescription(BuzillaBug bug, XmlLinkedNode xmlLinkedNode)
        {
            string fullDescription = string.Empty;
            foreach (XmlNode childNode in xmlLinkedNode.ChildNodes)
            {
                string actualDescription = string.Empty;

                switch (childNode.Name)
                {
                    case "bug_when":
                        bug.DescriptionWhen = childNode.InnerText;
                        break;
                    case "who":
                        bug.DescriptionWho = childNode.InnerText;
                        break;

                    case "thetext":
                        actualDescription = childNode.InnerText;
                        break;
                }
                //ToDo: take out const html strings
                //fullDescription = "<b><span style='font-size:8.5pt; font-family:\"Tahoma\",\"sans-serif\";background:silver;mso-fareast-font-family:\"Times New Roman\"'>";
                //.. is same as :-
                fullDescription = @"<b><span style='font-size:8.5pt; font-family:""Tahoma"",""sans-serif"";background:silver;mso-fareast-font-family:""Times New Roman""'>";

                if (bug.DescriptionWhen != null && bug.DescriptionWhen.Trim() != string.Empty)
                    fullDescription += DateTime.Parse(bug.DescriptionWhen).ToLongDateString() + " " + DateTime.Parse(bug.DescriptionWhen).ToLongTimeString();

                if (bug.DescriptionWho != null && bug.DescriptionWho.Trim() != string.Empty)
                    fullDescription += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + bug.DescriptionWho + ": ";
                fullDescription += "</span></b></br>";
                fullDescription += actualDescription;
                fullDescription += "</br></br>";
            }
            return fullDescription;
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
                            //WorkItemType wiType = store.Projects[3].WorkItemTypes[8]; //Project[3] => tfs2010Test, WorkItemTypes[8]=> BuzillaBug

                            var newWI = new WorkItem(workItemType);

                            newWI.Title = bug.Title;
                            
                            //ToDo: will be nice if done using attributes on properties and reflection

                            newWI.Fields["Company.Department.Bugzilla.ID"].Value = bug.bugzilla_BugId;

                            newWI.Fields["Company.Department.Bugzilla.ReportedOn"].Value = bug.CreatedDateTime;

                            newWI.Fields["Company.Department.Bugzilla.ModifiedDate"].Value = bug.ChangedDateTime;

                            newWI.Fields["Company.Department.Bugzilla.Product"].Value = bug.Product;

                            newWI.Fields["Company.Department.Bugzilla.Component"].Value = bug.Component;

                            newWI.Fields["Company.Department.Bugzilla.Version"].Value = bug.Version;

                            newWI.Fields["Company.Department.Bugzilla.Platform"].Value = bug.Platform;

                            newWI.Fields["Company.Department.Bugzilla.OS"].Value = bug.OperatingSystem;

                            //Note:             System.State wont allow values other than 'NEW' for first time because of the presence of Workflow(state transition) in place,
                            //Note continued:       Workitem template can not have empty workflow, workflow cannot start with multiple transition
                            //Note continued:       and System.State cannot copy other field values.
                            //Note continued:       So use 'Company.Department.Bugzilla.Status' field till we get all the Bugzilla bugs to TFS and then update the custom workitem to show System.Status
                            newWI.Fields["Company.Department.Bugzilla.Status"].Value = bug.BugStatus;

                            newWI.Fields["Company.Department.Bugzilla.Resolution"].Value = bug.Resolution;

                            newWI.Fields["Company.Department.Bugzilla.Priority"].Value = bug.Priority;

                            newWI.Fields["Company.Department.Bugzilla.BugSeverity"].Value = bug.Severity;

                            newWI.Fields["Company.Department.Bugzilla.TargetMilestone"].Value = bug.TargetMilestone;


                            if (bug.DependsOn != 0)
                            {
                                WorkItemCollection wicol =
                                    store.Query("Select [ID] from Issues where (Company.Department.Bugzilla.BugzillaId='" + bug.DependsOn +
                                                "')");
                                if (wicol.Count > 0)
                                {
                                    RelatedLink link = new RelatedLink(wicol[0].Id);
                                    newWI.Links.Add(link);
                                }
                            }

                            if (bug.Blocks != 0)
                            {
                                WorkItemCollection wicol =
                                    store.Query("Select [ID] from Issues where (Company.Department.Bugzilla.BugzillaId='" + bug.Blocks +
                                                "')");
                                if (wicol.Count > 0)
                                {
                                    RelatedLink link = new RelatedLink(wicol[0].Id);
                                    newWI.Links.Add(link);
                                }
                            }


                            newWI.Fields["Company.Department.Bugzilla.ReportedBy"].Value = bug.ReportedBy;

                            newWI.Fields["Company.Department.Bugzilla.AssignedTo"].Value = bug.AssignedTo;

                            if (bug.CcList != null)
                                newWI.Fields["Company.Department.Bugzilla.Observers"].Value = bug.CcList; // Company.Department.Bugzilla.EmailCCList

                            newWI.Fields["Company.Department.Bugzilla.Estimated_Time"].Value = bug.EstimatedTime;

                            newWI.Fields["Company.Department.Bugzilla.Remaining_Time"].Value = bug.RemainingTime;

                            newWI.Fields["Company.Department.Bugzilla.Actual_Time"].Value = bug.ActualTime;

                            if (bug.ExternalBugId != null)
                                newWI.Fields["Company.Department.Bugzilla.ExternalBugId"].Value = bug.ExternalBugId;

                            if (bug.ExternalReporter != null)
                                newWI.Fields["Company.Department.Bugzilla.ExternalReporter"].Value = bug.ExternalReporter;

                            if (bug.ReportedIn != null)
                                newWI.Fields["Company.Department.Bugzilla.ReportedIn"].Value = bug.ReportedIn;

                            if (bug.FixedIn != null)
                                newWI.Fields["Company.Department.Bugzilla.FixedIn"].Value = bug.FixedIn;

                            if (bug.DocImpact != null)
                                newWI.Fields["Company.Department.Bugzilla.DocImpact"].Value = bug.DocImpact;

                            if (bug.Attachments != null && bug.Attachments.Count > 0)
                            {
                                foreach (Attachment attachment in bug.Attachments)
                                {
                                    newWI.Attachments.Add(attachment);
                                }
                            }

                            //newWI.Description = bug.Description;

                            newWI.History = bug.Description;

                            //newWI.Fields["System.CreatedDate"].Value = bug.CreatedDateTime;
                            //newWI.Fields["System.CreatedBy"].Value = bug.ReportedBy;

                            newWI.Save();
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

        #endregion

        private void btnDeleteBugs_Click(object sender, RoutedEventArgs e)
        {
            int tfsBugStart, tfsBugEnd;

            if (!int.TryParse(txtTfsFrom.Text.Trim(), out tfsBugStart))
                return;

            if (!int.TryParse(txtTfsTo.Text.Trim(), out tfsBugEnd))
                return;

            tfsCol.Connect(ConnectOptions.None);

            var store = (WorkItemStore)tfsCol.GetService(typeof(WorkItemStore));

            //Note: The Project and WorkItemType has to be specifically selected in final implementation
#warning Select proper Project and WorkItemType

            IEnumerable<WorkItemOperationError> errors = store.Projects[3].Store.DestroyWorkItems(new[] { 6046, 6047, 6048, 6049, 6050, 6051, 6052, 6053, 6054, 6055 });

            if (errors != null)
            {
                MessageBox.Show("Delete failed. You don't have rights to delete the WIs.");
            }
            

        }


        #region Xml rpc Methods
        //private readonly Server _server;

        private void LoginClick(object sender, RoutedEventArgs e)
        {
            //if (!_server.LoggedIn)
            //    _server.Login("vishal.patil@Company.com", "vishal@bugzilla", true);

            //if (_server.LoggedIn)
            //    MessageBox.Show("You are now logged in");

            //HasBugzillaAccess = _server.LoggedIn;

        }

        private void Xml_Rpc_Click(object sender, RoutedEventArgs e)
        {
            //try
            //{
            //    string userEnteredIds = txtIds.Text.Trim();
            //    if (userEnteredIds == string.Empty)
            //    {    MessageBox.Show("Invalid Bug is");
            //        return;
            //    }
            //    bool isValidBugId = true;
            //    foreach (var userEnteredId in userEnteredIds.Split(",".ToCharArray()))
            //    {
            //        int id;
            //        if (int.TryParse(userEnteredId, out id))
            //        {
            //            //temporary unuse the return types                
            //            Bug bugs = _server.GetBug(id);
            //            BugComments detailedBug = _server.GetBugComments(id);
            //            BugAttachment attchment = _server.GetAttchment(id);
            //        }
            //        else
            //        {
            //            isValidBugId = false;
            //            MessageBox.Show("Some or all of the bug ids are not valid. Please correct and try again."); //" \n One that are valid are written at:" + ConfigurationManager.AppSettings.Get("OutputFileLocation"));
            //            break;
            //        }
            //    }
            //    if(isValidBugId)
            //    MessageBox.Show("Files are successfully written at :" + ConfigurationManager.AppSettings.Get("OutputFileLocation"));
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }

        private void LogoutClick(object sender, RoutedEventArgs e)
        {
            //if (_server.LoggedIn)
            //    _server.Logout();
            //HasBugzillaAccess = _server.LoggedIn;
        }

        public bool HasBugzillaAccess
        {
            get { return (bool)GetValue(HasBugzillaAccessProperty); }
            set { SetValue(HasBugzillaAccessProperty, value); }
        }

        // Using a DependencyProperty as the backing store for HasBuzillaAccess.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty HasBugzillaAccessProperty =
            DependencyProperty.Register("HasBugzillaAccess", typeof(bool), typeof(Window), new UIPropertyMetadata(false));


        #endregion


    }

}


