using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop;
using System.IO;
//https://github.com/matthewproctor/OutlookAttachmentExtractor/blob/master/Program.cs
namespace SDV_Outlook_Tools
{
    public partial class rb_sdv_outlook_tools
    {
        static int totalfilesize = 0;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btn_RemoveAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                RemoveAttachments();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler im Programmablauf", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RemoveAttachments()
        {
            try
            {
                DialogResult result1 = MessageBox.Show("Möchten Sie wirklich die Änhange aller Mails älterer als xx Tage entfernen?", "Entfernen der Änhänge", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                //Remove Attachments from Mails older than xx Days
                if (result1 == DialogResult.Yes)
                {
                    string pathToSave = getSaveFolder();
                    if (pathToSave != "0")
                    {
//                        Microsoft.Office.Interop.Outlook.Application Application = new Microsoft.Office.Interop.Outlook.Application();
//                        Microsoft.Office.Interop.Outlook.Accounts accounts = Application.Session.Accounts;
//                        foreach (Microsoft.Office.Interop.Outlook.Account account in accounts)
//                        {
//                            Console.WriteLine(account.DisplayName);
//                            //EnumerateFolders(account.Folder);
//                            EnumerateFoldersInDefaultStore(pathToSave);
//                        }
                        EnumerateFoldersInDefaultStore(pathToSave);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler im Programmablauf", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public static string getSaveFolder()
        {
            try
            {
                FolderBrowserDialog folderDlg = new FolderBrowserDialog
                {
                    ShowNewFolderButton = true
                };
                DialogResult result = folderDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    return folderDlg.SelectedPath;
                }
                else
                {
                    return "0";
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "0";
            }
        }

        static void EnumerateFoldersInDefaultStore(string pathToSaveFile)
        {
            Microsoft.Office.Interop.Outlook.Application Application = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Microsoft.Office.Interop.Outlook.Folder;
            EnumerateFolders(root, pathToSaveFile);
        }

        static void EnumerateFolders(Microsoft.Office.Interop.Outlook.Folder folder, string pathToSaveFile)
        {
            Microsoft.Office.Interop.Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Microsoft.Office.Interop.Outlook.Folder childFolder in childFolders)
                {
                    // We only want Inbox folders - ignore Contacts and others
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        MessageBox.Show(childFolder.FolderPath);
                        // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                        EnumerateFolders(childFolder, pathToSaveFile);
                    }
                }
            }
            MessageBox.Show("Checking in " + folder.FolderPath);
            IterateMessages(folder, pathToSaveFile);
        }

        static void IterateMessages(Microsoft.Office.Interop.Outlook.Folder folder, string basePath)
        {
            var fi = folder.Items.Restrict("[Unread] = true");
            if (fi != null)
            {
                foreach (Object item in fi)
                {
                    Microsoft.Office.Interop.Outlook.MailItem mi = (Microsoft.Office.Interop.Outlook.MailItem)item;
                    var attachments = mi.Attachments;
                    if (attachments.Count != 0)
                    {
                        // Create a directory to store the attachment 
                        if (!Directory.Exists(basePath + folder.FolderPath))
                        {
                            Directory.CreateDirectory(basePath + folder.FolderPath);
                        }

                        for (int i = 1; i <= mi.Attachments.Count; i++)
                        {
                            var fn = mi.Attachments[i].FileName.ToLower();
                            //check wither any of the strings in the extensionsArray are contained within the filename
//                            if (extensionsArray.Any(fn.Contains))
//                            {

                                // Create a further sub-folder for the sender
                                if (!Directory.Exists(basePath + folder.FolderPath + @"\" + mi.Sender.Address))
                                {
                                    Directory.CreateDirectory(basePath + folder.FolderPath + @"\" + mi.Sender.Address);
                                }
                                //totalfilesize = totalfilesize + mi.Attachments[i].Size;
                                if (!File.Exists(basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[i].FileName))
                                {
                                    Console.WriteLine("Saving " + mi.Attachments[i].FileName);
                                    mi.Attachments[i].SaveAsFile(basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[i].FileName);
                                    mi.Body = mi.Body + "Anhange nach " + basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[i].FileName+" gespeichert.";
                                   //mi.Attachments[i].Delete();
                                }
                                else
                                {
                                    Console.WriteLine("Already saved " + mi.Attachments[i].FileName);
                                }
  //                          }






                                //mi.Attachments[i].SaveAsFile(pathToSaveFile);
                                //mi.Attachments[i].Delete();
                                Console.WriteLine("Subject: " +mi.Subject+" Attachment: " + mi.Attachments[i].FileName);
                        }
                    }
                }
            }
        }
    }
}
