/*
 This program downloaded from https://gallery.technet.microsoft.com/Alternate-to-Timer-Job-d840eb2b/view/Discussions
 */
using ExportListToExcel.Properties;
using SP = Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;
using System.Data;
using System.Reflection;

namespace ExportListToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (SP.ClientContext clientContext = new SP.ClientContext(Settings.Default.SiteUrl))
            {
                string pwd = Settings.Default.password;
                SecureString pwdSecure = new SecureString();
                foreach (char c in pwd.ToCharArray())
                    pwdSecure.AppendChar(c);

                if (Settings.Default.SPOnline)
                {
                    clientContext.Credentials = new SP.SharePointOnlineCredentials(Settings.Default.username, pwdSecure);
                }
                else
                {
                    clientContext.Credentials = new NetworkCredential(Settings.Default.username, pwdSecure, Settings.Default.domain);
                }

                try
                {
                    DirectoryInfo dir = new DirectoryInfo(Settings.Default.Directory);
                    dir.Create();
                    string excelFileName = string.Format(Path.Combine(Settings.Default.Directory,"ListItem_{0}_{1}.xls"), Settings.Default.ListName, DateTime.Now.Ticks.ToString());

                    FileInfo file = new FileInfo(excelFileName);
                    StreamWriter streamWriter = file.CreateText();

                    StringWriter stringWriter = new StringWriter();
                    HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);

                    SP.List myList = clientContext.Web.Lists.GetByTitle(Settings.Default.ListName);
                    SP.View listView = myList.Views.GetByTitle("All Items");

                    Table tblListView = new Table();
                    tblListView.ID = "_tblListView";
                    tblListView.BorderStyle = BorderStyle.Solid;
                    tblListView.BorderWidth = Unit.Pixel(1);
                    tblListView.BorderColor = Color.Silver;

                    listView.RowLimit = 2147483647;

                    Console.WriteLine("Load listView");

                    clientContext.Load(listView);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("listView loaded and query executed.");

                    SP.CamlQuery query = new SP.CamlQuery();
                    query.ViewXml = "<View><Query>" + listView.ListViewXml + "</Query></View>";
                    
                    SP.ListItemCollection items = myList.GetItems(query);
                    clientContext.Load(myList);
                    clientContext.Load(myList.Fields);
                    clientContext.Load(listView.ViewFields);
                    clientContext.Load(items);

                    Console.WriteLine("myList, myList.Fields, listView.ViewFields, items loaded");
                    
                    clientContext.ExecuteQuery();

                    Console.WriteLine("query executed.");

                    if (items != null && items.Count > 0)
                    {
                        DataTable dt = new DataTable();
                        foreach (var field in items[0].FieldValues.Keys)
                        {
                            dt.Columns.Add(field);
                        }

                        foreach (var item in items)
                        {
                            DataRow dr = dt.NewRow();

                            foreach (var obj in item.FieldValues)
                            {
                                if (obj.Value != null)
                                {
                                    string type = obj.Value.GetType().FullName;

                                    if (type == "Microsoft.SharePoint.Client.FieldLookupValue")
                                    {
                                        dr[obj.Key] = ((SP.FieldLookupValue)obj.Value).LookupValue;
                                    }
                                    else if (type == "Microsoft.SharePoint.Client.FieldUserValue")
                                    {
                                        dr[obj.Key] = ((SP.FieldUserValue)obj.Value).LookupValue;
                                    }
                                    else
                                    {
                                        dr[obj.Key] = obj.Value;
                                    }
                                }
                                else
                                {
                                    dr[obj.Key] = null;
                                }
                            }

                            dt.Rows.Add(dr);
                        }




                        DataView dvListViewData = dt.DefaultView;
                        if (dvListViewData != null && dvListViewData.Count > 0)
                        {
                            tblListView.Rows.Add(new TableRow());
                            tblListView.Rows[0].BackColor = Color.Gainsboro;
                            tblListView.Rows[0].Font.Bold = true;

                            for (int i = 0; i < listView.ViewFields.Count; i++)
                            {
                                tblListView.Rows[0].Cells.Add(new TableCell());
                                tblListView.Rows[0].Cells[i].Text = listView.ViewFields[i].ToString().Replace("LinkTitle","Title");
                            }

                            for (int i = 0; i < dvListViewData.Count; i++)
                            {
                                tblListView.Rows.Add(new TableRow());

                                for (int j = 0; j < listView.ViewFields.Count; j++)
                                {
                                    tblListView.Rows[i + 1].Cells.Add(new TableCell());

                                    if (dt.Columns.Contains(listView.ViewFields[j].ToString().Replace("LinkTitle", "Title")))
                                    {
                                        tblListView.Rows[i + 1].Cells[j].Text = dvListViewData[i][listView.ViewFields[j].ToString().Replace("LinkTitle","Title")].ToString();
                                    }
                                }
                            }
                        }
                    }




                    tblListView.RenderControl(htmlTextWriter);
                    streamWriter.Write(stringWriter.ToString());

                    htmlTextWriter.Close();
                    streamWriter.Close();
                    stringWriter.Close();

                    Console.WriteLine("completed");
                }                   
                catch (Exception ex)
                {
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Authentication failed." + ex.Message);
                    Console.ForegroundColor = ConsoleColor.Gray;

            
                }
            }

            
        }




        
        
    }    
}
