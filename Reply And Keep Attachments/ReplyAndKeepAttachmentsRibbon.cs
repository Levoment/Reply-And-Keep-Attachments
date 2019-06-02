using System.IO;
using System.Windows;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace Reply_And_Keep_Attachments
{
    public partial class ReplyAndKeepAttachmentsRibbon
    {
        // Variable for controlling whether the reply all button gets clicked
        private bool replyAll = false;

        private void ReplyAndKeepAttachmentsRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ReplyAndKeepAttachmentsButton_Click(object sender, RibbonControlEventArgs e)
        {
            // Set the replyAll variable
            replyAll = false;
            replyAndKeepAttachments(sender, e);
        }

        private void replyAndKeepAttachments(object sender, RibbonControlEventArgs e)
        {
            // Get the inspector
            Inspector inspector = e.Control.Context as Inspector;
            // Get the current mail message
            MailItem mailItem = inspector.CurrentItem as MailItem;

            // Check if we have a mail message to work with
            if (mailItem != null)
            {
                // Check if we have attachments
                if (mailItem.Attachments.Count > 0)
                {
                    // Variable holding the reply
                    MailItem mailItemToSend;

                    // Check if Reply All button was pressed
                    if (replyAll)
                    {
                        // Create the reply all message item
                        mailItemToSend = mailItem.ReplyAll();
                    }
                    else
                    {
                        // Create the reply message item
                        mailItemToSend = mailItem.Reply();
                    }

                    // Iterate through each attachment
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        try
                        {
                            string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";
                            // Get property of attachment (ATT_INVISIBLE_IN_HTML | ATT_INVISIBLE_IN_RTF)
                            var attachmentFlags = attachment.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS);
                            // If attachment is not ATT_MHTML_REF
                            if (attachmentFlags != 4)
                            {
                                // Path of the folder to save files before attaching
                                string folderPath = @"C:\ReplyKeepAttachmentsTmp";
                                try
                                {
                                    // Check if the folder does not exist
                                    if (!Directory.Exists(folderPath))
                                    {
                                        // Create the folder if it doesn't exist
                                        DirectoryInfo directoryInfo = Directory.CreateDirectory(folderPath);
                                    }

                                    // Check if the attachment is not an embeeded element
                                    if ((int)attachment.Type != 6)
                                    {
                                        // Save current attachment
                                        attachment.SaveAsFile(Path.Combine(@"C:\ReplyKeepAttachmentsTmp\", attachment.FileName));
                                        if (mailItemToSend != null)
                                        {
                                            // Attach it to the reply message
                                            mailItemToSend.Attachments.Add(Path.Combine(@"C:\ReplyKeepAttachmentsTmp\", attachment.FileName));
                                            //mailItemToSend.Save();
                                            // Delete the temporary file
                                            if (File.Exists(Path.Combine(@"C:\ReplyKeepAttachmentsTmp\", attachment.FileName)))
                                            {
                                                File.Delete(Path.Combine(@"C:\ReplyKeepAttachmentsTmp\", attachment.FileName));
                                            }
                                        }
                                    }

                                }
                                catch (System.IO.IOException ioException)
                                {
                                    // Show a message showing that there was an error while attempting to create the temporary folder
                                    MessageBox.Show("There was an error while attempting to create the temporary folder to save attachments:\n\n" +
                                        ioException.Message, "Reply and Keep Attachments Add-In Message", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                catch (System.Exception genericException)
                                {
                                    // Show the exception message
                                    MessageBox.Show(genericException.Message, "Reply and Keep Attachments Add-In Message", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                finally { }
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException comException)
                        {
                            MessageBox.Show(comException.Message, "Reply and Keep Attachments Add-In Message", MessageBoxButton.OK, MessageBoxImage.Error);
                        }

                    }
                    // Check if the message to send is valid
                    if (mailItemToSend != null)
                    {
                        // Open the reply Window to allow the user to edit and send the message
                        mailItemToSend.GetInspector.Activate();
                    }
                }
                else
                {
                    // Show a message saying that no attachments could be found
                    MessageBox.Show("No attachments could be found on this mail message.", "Reply and Keep Attachments Add-In Message", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
        }

        private void ReplyAllAndKeepAttachmentsButton_Click(object sender, RibbonControlEventArgs e)
        {
            // Set the replyAll variable
            replyAll = true;
            replyAndKeepAttachments(sender, e);
        }
    }
}
