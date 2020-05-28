using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.SharePoint.Client;
using System.Security;

namespace CsomWpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        

        public MainWindow()
        {
            InitializeComponent();
            this.show();
        }

        public void show()
        {
            string UserName = "omisayeduiu@omisayeduiu.onmicrosoft.com", Password = "Omi01683539680";
            ClientContext ctx = new ClientContext("https://omisayeduiu.sharepoint.com/sites/sayeddev");
            SecureString passWord = new SecureString();
            foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
            ctx.Credentials = new SharePointOnlineCredentials(UserName, passWord);

            Web myWeb = ctx.Web;
            Microsoft.SharePoint.Client.List sectionsList = myWeb.Lists.GetByTitle("sectionInfo");

            CamlQuery myQuery = new CamlQuery();
            myQuery.ViewXml = @"<view>
 
 
<Query>
   <Where>
      <Or>
         <And>
            <Eq>
               <FieldRef Name='sectionName' />
               <Value Type='Text'>A</Value>
            </Eq>
            <Geq>
               <FieldRef Name='sectionWeight' />
               <Value Type='Number'>21</Value>
            </Geq>
         </And>
         <And>
            <Eq>
               <FieldRef Name='sectionName' />
               <Value Type='Text'>B</Value>
            </Eq>
            <Lt>
               <FieldRef Name='sectionWeight' />
               <Value Type='Number'>26</Value>
            </Lt>
         </And>
      </Or>
   </Where>
</Query>
<ViewFields>
   <FieldRef Name='ID' />
   <FieldRef Name='sectionName' />
   <FieldRef Name='sectionWeight' />
</ViewFields>
<QueryOptions />


                                </view>";

            Microsoft.SharePoint.Client.ListItemCollection mySections = sectionsList.GetItems(myQuery);
            ctx.Load(mySections);
            ctx.ExecuteQuery();
            string result = string.Empty;
            result = "ID" + " " + "Section" + " " + "Weight" + Environment.NewLine;
            foreach (Microsoft.SharePoint.Client.ListItem itm in mySections)
            {

                result += itm["ID"].ToString() + "     " + itm["sectionName"].ToString() + "         " + itm["sectionWeight"].ToString() + Environment.NewLine;

            }
            showBox.Text = result;


        }
       
        private void button_Click(object sender, RoutedEventArgs e)
        {

            this.show();
        }

        private void createButton_Click(object sender, RoutedEventArgs e)
        {
            string UserName = "omisayeduiu@omisayeduiu.onmicrosoft.com", Password = "Omi01683539680";
            ClientContext ctx = new ClientContext("https://omisayeduiu.sharepoint.com/sites/sayeddev");
            SecureString passWord = new SecureString();
            foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
            ctx.Credentials = new SharePointOnlineCredentials(UserName, passWord);


            Web myWeb = ctx.Web;
            Microsoft.SharePoint.Client.List sectionsList = myWeb.Lists.GetByTitle("sectionInfo");

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            Microsoft.SharePoint.Client.ListItem newItem = sectionsList.AddItem(itemCreateInfo);
            newItem["sectionName"] = sectionBox.Text.Trim().ToString();
            newItem["sectionWeight"] =Convert.ToInt32(weightBox.Text.Trim());
            newItem.Update();

            ctx.ExecuteQuery();

            this.show();
        }

        private void updateButton_Click(object sender, RoutedEventArgs e)
        {
            string UserName = "omisayeduiu@omisayeduiu.onmicrosoft.com", Password = "Omi01683539680";
            ClientContext ctx = new ClientContext("https://omisayeduiu.sharepoint.com/sites/sayeddev");
            SecureString passWord = new SecureString();
            foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
            ctx.Credentials = new SharePointOnlineCredentials(UserName, passWord);

            Web myWeb = ctx.Web;
            Microsoft.SharePoint.Client.List sectionsList = myWeb.Lists.GetByTitle("sectionInfo");
            Microsoft.SharePoint.Client.ListItem item = sectionsList.GetItemById(Convert.ToInt32(updateId.Text.Trim()));
            item["sectionName"] = updateSection.Text.Trim();
            item["sectionWeight"] = Convert.ToInt32(updateWeight.Text.Trim());
            item.Update();
            ctx.ExecuteQuery();
            this.show();
        }

        private void deleteButton_Click(object sender, RoutedEventArgs e)
        {
            string UserName = "omisayeduiu@omisayeduiu.onmicrosoft.com", Password = "Omi01683539680";
            ClientContext ctx = new ClientContext("https://omisayeduiu.sharepoint.com/sites/sayeddev");
            SecureString passWord = new SecureString();
            foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
            ctx.Credentials = new SharePointOnlineCredentials(UserName, passWord);

            Web myWeb = ctx.Web;
            Microsoft.SharePoint.Client.List sectionsList = myWeb.Lists.GetByTitle("sectionInfo");

            Microsoft.SharePoint.Client.ListItem item = sectionsList.GetItemById(Convert.ToInt32(deleteId.Text.Trim()));
            item.DeleteObject();
            ctx.ExecuteQuery();
            this.show();
        }
    }
}
