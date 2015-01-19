using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Client;
using System.Configuration;

namespace xmlRpcExample
{
    /// <summary>
    /// Interaction logic for DeleteWorkItem.xaml
    /// </summary>
    public partial class DeleteWorkItem : Window
    {
        private TfsTeamProjectCollection from_tfsCol;
        WorkItemStore _from_store;

        public DeleteWorkItem()
        {
            InitializeComponent();

            var tfsUriSetting = ConfigurationManager.AppSettings.Get("Delete_TFSUri");
            var tfsUri = new Uri(tfsUriSetting);
            from_tfsCol = new TfsTeamProjectCollection(tfsUri);
            _from_store = (WorkItemStore)from_tfsCol.GetService(typeof(WorkItemStore));

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int workItemId = int.Parse(txtWorkItemId.Text.Trim());
            string from_projectName = ConfigurationManager.AppSettings["Delete_ProjectName"];
             var decide =  MessageBox.Show("Work Item with id: " + workItemId + " will be deleted from Project: " + from_tfsCol.Name + "-> " + from_projectName + " \n Proceed?", "Delete", MessageBoxButton.YesNo);

             if (decide == MessageBoxResult.No) return;

             DistroyWorkItem(workItemId);
        }

        private void DistroyWorkItem(int workItemId)
        {
          var error=  _from_store.DestroyWorkItems(new List<int> { workItemId });
          foreach (var item in error)
          {
              var xx = item;
          }
        }
    }
}
