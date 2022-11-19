using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Windows.Forms;
using ExcelUtil;
using JazzApp;
using Ftp;
using System.Collections;

namespace SendNewsLetter
{
    /// <summary> Main class for the SendNewsLetter application</summary>
    public class NewsletterMain
    {
        #region FTP passwords

        /// <summary>FTP password for jazzliv1</summary>
        private string m_ftp_password = "TODO";

        #endregion // FTP passwords

        /// <summary>Initialize XML objects
        /// <para>XML objects for JazzApplication.xml and current season XML file will be created</para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        public bool InitXml(out string o_error)
        {
            o_error = @"";

            return NewsLetterXml.Init(m_ftp_password, FileUtil.GetPathToExeDirectory(), out o_error);

        } // InitXml

        /// <summary>Download new versions, posters, attachment files and templates from the server
        /// <para>1. Check that Internet connection is available. Call of JazzFtp.Execute.IsConnectionAvailable</para>
        /// <para>2. Create an instance of Ftp.DownLoad. Input data is host, user and password</para>
        /// <para>3. Set local directory name (call of FileUtil.PosterDirectory) and server name (hardcoded) for the posters</para>
        /// <para>4. Get the poster files. Call of Ftp.DownLoad.GetFiles</para>
        /// <para>5. Set local directory name (call of FileUtil.AttachmentDirectory) and server name (hardcoded) for the attachments</para>
        /// <para>6. Get the attachment files. Call of Ftp.DownLoad.GetFiles</para>
        /// <para>7. If i_new_version_templates equal to true:</para>
        /// <para>7.1 - Set local directory name (call of FileUtil.NewVersionDirectory) and server name (hardcoded) for the installer</para>
        /// <para>7.2 - Get the attachment files. Call of Ftp.DownLoad.GetFiles</para>
        /// <para>7.3 - Set local directory name (call of FileUtil.TemplatesDirectory) and server name (hardcoded) for the templates</para>
        /// <para>7.4 - Get the template files. Call of Ftp.DownLoad.GetFiles</para>
        /// <para></para>
        /// <para>Posters and attachchments or the new version is downloaded (Version 1.10)</para>
        /// <para>Reason: If a poster is displayed in the application it can't be downloaded.</para>
        /// </summary>
        /// <param name="i_new_version_templates">Flag telling if an installer for a new version shall be downloaded</param>
        /// <param name="o_error">Error message</param>
        public bool DownloadFiles(bool i_new_version_templates, out string o_error)
        {
            o_error = "";

            string status_message = @"";
            string path_exe_dir = FileUtil.GetPathToExeDirectory();
            if(!JazzFtp.Execute.IsConnectionAvailable(path_exe_dir, out status_message))
            {
                o_error = NewsLetterSettings.Default.ErrMsgNoInternetConnection;
                return false;
            }

            bool b_down_load = true;

            Ftp.DownLoad ftp_download = new Ftp.DownLoad(NewsLetterSettings.Default.FtpHost, NewsLetterSettings.Default.FtpUser, m_ftp_password);

            string server_address_directory = @"";
            string local_address_directory = @"";

            if (!i_new_version_templates)
            {
                server_address_directory = @"/setup/Newsletter/Plakat/";
                local_address_directory = FileUtil.PosterDirectory();

                if (!ftp_download.GetFiles(server_address_directory, local_address_directory, out o_error))
                {
                    b_down_load = false;
                }

                server_address_directory = @"/setup/Newsletter/Anhang/";
                local_address_directory = FileUtil.AttachmentDirectory();

                if (!ftp_download.GetFiles(server_address_directory, local_address_directory, out o_error))
                {
                    b_down_load = false;
                }
            }

            if (i_new_version_templates)
            {
                server_address_directory = @"/setup/Newsletter/NeueVersion/";
                local_address_directory = FileUtil.NewVersionDirectory();

                if (!ftp_download.GetFiles(server_address_directory, local_address_directory, out o_error))
                {
                    b_down_load = false;
                }

                server_address_directory = @"/setup/Newsletter/Vorlagen/";
                local_address_directory = FileUtil.TemplatesDirectory();

                if (!ftp_download.GetFiles(server_address_directory, local_address_directory, out o_error))
                {
                    b_down_load = false;
                }
            }

            if (!b_down_load)
            {
                o_error = NewsLetterSettings.Default.ErrMsgFilesDownload;
            }

            return b_down_load;
        }

        /// <summary>Get downloaded picture names
        /// <para>1. Check that Internet connection is available. Call of JazzFtp.Execute.IsConnectionAvailable</para>
        /// <para>2. Create an instance of Ftp.DownLoad. Input data is host, user and password</para>
        /// <para>3. Set local directory name (call of FileUtil.PosterDirectory) and server name (hardcoded) for the posters</para>
        /// <para>4. Get the poster files. Call of Ftp.DownLoad.GetFileNames</para>
        /// </summary>
        /// <param name="o_picture_names_array">Array of downloaded picture file names</param>
        /// <param name="o_error">Error message</param>
        public bool GetDownloadedPictureNames(out string[] o_picture_names_array, out string o_error)
        {
            o_error = "";
            ArrayList picture_names_array = new ArrayList();
            o_picture_names_array = (string[])picture_names_array.ToArray(typeof(string));

            string status_message = @"";
            string path_exe_dir = FileUtil.GetPathToExeDirectory();
            if (!JazzFtp.Execute.IsConnectionAvailable(path_exe_dir, out status_message))
            {
                o_error = NewsLetterSettings.Default.ErrMsgNoInternetConnection;
                return false;
            }

            bool b_down_load = true;

            Ftp.DownLoad ftp_download = new Ftp.DownLoad(NewsLetterSettings.Default.FtpHost, NewsLetterSettings.Default.FtpUser, m_ftp_password);

            // string server_address_directory = @"/Newsletter/Plakat/";
            string server_address_directory = @"/setup/Newsletter/Plakat/";
            string local_address_directory = FileUtil.PosterDirectory();

            if (!ftp_download.GetFileNames(server_address_directory, local_address_directory, out o_picture_names_array, out o_error))
            {
                b_down_load = false;
            }

            return b_down_load;
        }

        /// <summary>Download the file with addresses from the Web server
        /// <para>1. Create an instance of Ftp.DownLoad. Input data is host, user and password</para>
        /// <para>2. Check that Internet connection is available. Call of JazzFtp.Execute.IsConnectionAvailable</para>
        /// <para>3. Get server name for the address file from the configuration file (NewsLetterSettings)</para>
        /// <para>4. Get local name for the address file. Call of FileUtil.InputCsvFileName</para>
        /// <para>5. Download the address file. Call of Ftp.DownLoad.GetFile</para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        public bool DownloadAddresseFile(out string o_error)
        {
            o_error = "";

            Ftp.DownLoad ftp_download = new Ftp.DownLoad(NewsLetterSettings.Default.FtpHost, NewsLetterSettings.Default.FtpUser, m_ftp_password);

            string status_message = @"";
            string path_exe_dir = FileUtil.GetPathToExeDirectory();
            if (!JazzFtp.Execute.IsConnectionAvailable(path_exe_dir, out status_message))
            {
                o_error = NewsLetterSettings.Default.ErrMsgNoInternetConnection;
                return false;
            }

            string server_address_file = "/adressen/" + SendNewsLetter.NewsLetterSettings.Default.FileExcel;

            string local_address_file = FileUtil.InputCsvFileName();

            if (!ftp_download.GetFile(server_address_file, local_address_file, null, out o_error))
            {
                o_error = NewsLetterSettings.Default.ErrMsgNoExcelFileDownload;
                return false;
            }

            return true;
        }

        /// <summary>Main function that sends test newsletter(s) or newsletters to all in the .csv file that has subscribed to the newsletter
        /// <para>1. Check Email input data i_newsletter_data. Call of NewsletterData.Check</para>
        /// <para>2. Get the body HTML text as string. Call of _BodyHtml</para>
        /// <para>3. Get the body PLAIN text as string. Call of _BodyText</para>
        /// <para>4. Get Email addresses if i_only_to is false. Call of _GetEmailAdresses.</para>
        /// <para>5. Create HTLM and PLAIN AlternateView. Calls of AlternateView.CreateAlternateViewFromString</para>
        /// <para>5. Get poster file name. Call of NewsletterData.GetPosterName</para>
        /// <para>6. Add poster to HTML view. Call of _AddPosterToHtmlView</para>
        /// <para>7. Create MailMessage instance. Input data is to, from, subject and body text</para>
        /// <para>8. Add poster as attachment. Call of _AddPosterAsAttachment</para>
        /// <para>9. Add views HTML and PLAIN body to MailMessage. Calls of MailMessage.AlternateViews.Add</para>
        /// <para>10. Send Emails. Call of _SendEmails</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Emails that shall be sent</param>
        ///  <param name="i_test_send">Flag telling if only test Email(s) shall be sent</param>
        ///  <param name="i_only_to">Flag telling if an Email shall be sent to only one person</param>
        ///  <param name="o_error">Error message</param>
        private bool _SendNewsLetter(NewsletterData i_newsletter_data, bool i_test_send, bool i_only_to, out string o_error)
        {
            o_error = "";

            if (!i_newsletter_data.Check(out o_error)) return false;

            string str_body_html = _BodyHtml(i_newsletter_data);
            string str_body_text = _BodyText(i_newsletter_data);

            string email_adresses = "";
            Stack<string> bcc_email_adresses = new Stack<string>();
            if (!i_only_to)
            {
                if (!_GetEmailAdresses(i_newsletter_data, i_test_send, out email_adresses, ref bcc_email_adresses, out o_error)) return false;
            }

            string poster_file_name = i_newsletter_data.GetPosterName();

            AlternateView html_view = AlternateView.CreateAlternateViewFromString(str_body_html, null, "text/html");

            AlternateView plain_view = AlternateView.CreateAlternateViewFromString(str_body_text, null, "text/plain");

            _AddPosterToHtmlView(poster_file_name, ref html_view);

            MailMessage mail_poster = new MailMessage(i_newsletter_data.m_from, i_newsletter_data.m_to, i_newsletter_data.m_subject, str_body_text);

            _AddPosterAsAttachment(poster_file_name, ref mail_poster);

            _AddAttachment(i_newsletter_data.GetAttachmentName(), ref mail_poster);
            
            mail_poster.AlternateViews.Add(html_view);
            // mail_poster.AlternateViews.Add(plain_view);

            if (!_SendEmails(i_newsletter_data, bcc_email_adresses, i_test_send, ref mail_poster, out o_error)) return false;


            return true;
        }

        /// <summary>Create mail and send emails
        /// <para>1.Create an instance of SmtpClient. Input data is server name from i_newsletter_data</para>
        /// <para>2.Create an instance of NetworkCredential and add to SmtpClient. Call of SmtpClient.Credentials</para>
        /// <para>3. Set port 587 in SmtpClient. Call of SmtpClient.Port</para>
        /// <para>4. Set batch size to 60 Emails. The Internet provider has a limit (100 for Hostpoint)</para>
        /// <para>5. If i_bcc_email_adresses is empty: Call _SendOneEmail and return.</para>
        /// <para>6. Loop over all Email addresses in i_bcc_email_adresses</para>
        /// <para>6.1 - Create instance of MailAddress. Add Email adresses until 60 addresses.Calls of MailAddress.Bcc.Add</para>
        /// <para>6.2 - Send 60 Emails (or less). Call of _SendBatch</para>
        /// <para>6.3 - End of loop (break) if _SendBatch returns cancel flag equal to true (input from the user)</para>
        /// <para>6.4 - Message box for the (test) case that Emailletters not shall be sent</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Emails that shall be sent</param>
        ///  <param name="i_bcc_email_adresses">Array of Email addresses</param>
        ///  <param name="i_test_send">Flag telling if only test Email(s) shall be sent</param>
        ///   <param name="i_mail_poster">MailMessage instance reference (pointer)</param>
        ///  <param name="o_error">Error message</param>
        private bool _SendEmails(NewsletterData i_newsletter_data, Stack<string> i_bcc_email_adresses, bool i_test_send, ref MailMessage i_mail_poster, out string o_error)
        {
            o_error = "";
            SmtpClient smtp_client = new SmtpClient(i_newsletter_data.m_server_name);
            smtp_client.Credentials = new NetworkCredential(i_newsletter_data.m_credential_email, i_newsletter_data.m_credential_passw);
            smtp_client.Port = 587; // Default port 25 cannot be used in Sweden. The Internet provider Hostpoint solved this problem

            int n_bcc_batch = 60;
            int i_bcc_batch = 0;
            int n_bcc_total = i_bcc_email_adresses.Count;
            int i_bcc_total = 0;
            i_mail_poster.Bcc.Clear();

            if (0 == i_bcc_email_adresses.Count)
            {
                _SendOneEmail(i_mail_poster, smtp_client, out o_error);

                return true;
            }

            bool o_cancel = false;
            foreach (string str_bcc in i_bcc_email_adresses)
            {
                i_bcc_total = i_bcc_total + 1;
                i_bcc_batch = i_bcc_batch + 1;
                // _LogfileAppend("  m_bcc_stack: " + str_bcc, log_file_name);
                MailAddress adress_bcc = new MailAddress(str_bcc);
                i_mail_poster.Bcc.Add(adress_bcc);

                if (n_bcc_batch == i_bcc_batch || n_bcc_total == i_bcc_total)
                {
                    if (!_SendBatch(i_mail_poster, smtp_client, i_test_send, ref o_cancel, out o_error)) return false;

                    if (o_cancel)
                    {
                        break;
                    }

                    i_bcc_batch = 0;
                    i_mail_poster.Bcc.Clear();

                    if (n_bcc_total == i_bcc_total)
                    {
                        break;
                    }
                }
            }

            if (!o_cancel)
            {
                if (!NewsLetterSettings.Default.DoNotSendAll)
                {
                    string error_message = @"";

                    if (!i_test_send) // Only for the case all emails 
                    {
                        if (!NewsletterSave.Execute(i_newsletter_data, out error_message))
                        {
                            MessageBox.Show(error_message, " ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            return false;
                        }
                    }

                    string msg_box_text = NewsLetterSettings.Default.EndSendAllMessage;
                    if (i_test_send)
                    {
                        // Start temporarely QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ
                        /*
                        if (!NewsletterSave.Execute(i_newsletter_data, out error_message))
                        {
                            MessageBox.Show(error_message, " ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            return false;
                        }

                        MessageBox.Show("The Homepage -> Archive -> Newsletter was updatet for this test-version of Newsletter." + 
                            " In the production-version the homepage will only be updated when all newsletters are sent.");
                        */
                        // End temporarely QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ

                        msg_box_text = NewsLetterSettings.Default.EndSendTestMessage;
                    }
                    DialogResult end_result =
                        MessageBox.Show(msg_box_text, "Succesful",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);

                } // DoNotSendAll

            } // o_cancel = false

            return true;
        }

        /// <summary>Send batch (60 or less) of Emails
        /// <para>1. Let the user confirm that the Emails shall be sent. Return with flag o_cancel equal to false if cancelled.</para>
        /// <para>2. Show message for the test case that Emails not shall be sent (DoNotSendAll in configuration file) and return.</para>
        /// <para>3. Send Emails. Call of SmtpClient Send with i_mail_poster as input data</para>
        /// </summary>
        ///  <param name="i_mail_poster">Instance of MailMessage that hold all data for the Emails</param>
        ///  <param name="i_smtp_client">SMTP client data</param>
        ///  <param name="i_test_send">Flag telling if only test Email(s) shall be sent</param>
        ///   <param name="o_cancel">Flag telling if the user cancelled the sending of Emails</param>
        ///  <param name="o_error">Error message</param>
        private bool _SendBatch(MailMessage i_mail_poster, SmtpClient i_smtp_client, bool i_test_send, ref bool o_cancel, out string o_error)
        {
            o_cancel = false;
            o_error = "";

            DialogResult confirm_result =
                MessageBox.Show(NewsLetterSettings.Default.BatchSendMessage + " \n"
                + i_mail_poster.Bcc.ToString() + "\n", "Succesful",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
            if (DialogResult.OK != confirm_result)
            {
                o_cancel = true;

                return true;
            }

            if (!i_test_send && NewsLetterSettings.Default.DoNotSendAll)
            {
                MessageBox.Show(NewsLetterSettings.Default.TestSendMessage, "Test version of application Newsletter",
                     MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return true;
            }

            try
            {
                i_smtp_client.Send(i_mail_poster);
            }
            catch (SmtpFailedRecipientException smtpEx)
            {
                o_error = "Recipient Exception: " + smtpEx.ToString();

                MessageBox.Show("Recipient Exception: " + smtpEx.ToString(), "Send Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                MessageBox.Show("Recipients: " + i_mail_poster.Bcc.ToString(), "Send Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return true; 
            }
            catch (Exception ex)
            {
                o_error = "Exception from Send: " + ex.ToString();

                MessageBox.Show("Failure sending eMails: " + ex.ToString(), "Send Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        /// <summary>Send one Email</summary>
        ///  <param name="i_mail_poster">Instance of MailMessage that hold all data for the Emails</param>
        ///  <param name="i_smtp_client">SMTP client data</param>
        ///  <param name="o_error">Error message</param>
        private bool _SendOneEmail(MailMessage i_mail_poster, SmtpClient i_smtp_client, out string o_error)
        {
            o_error = "";

            i_mail_poster.Bcc.Clear();

            try
            {
                i_smtp_client.Send(i_mail_poster);
            }
            catch (Exception ex)
            {
                o_error = "Exception from _SendOneEmail: " + ex.ToString();

                MessageBox.Show("Failure sending Email: " + ex.ToString(), "Send Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        /// <summary>Add link to poster</summary>
        private void _AddPosterToHtmlView(string i_poster_file_name, ref AlternateView io_html_view)
        {
            if (i_poster_file_name == "" || !SendNewsLetter.NewsLetterSettings.Default.PosterEmbedded)
            {
                return;
            }

            LinkedResource poster = new LinkedResource(i_poster_file_name);
            poster.ContentId = "plakat";
            io_html_view.LinkedResources.Add(poster);
        }

        /// <summary>Add poster as attachement</summary>
        private void _AddPosterAsAttachment(string i_poster_file_name, ref MailMessage o_mail_poster)
        {
            if (i_poster_file_name == "" || !SendNewsLetter.NewsLetterSettings.Default.PosterAttached)
            {
                return;
            }

            if (!File.Exists(i_poster_file_name))
            {
                return;
            }

            Attachment data = new Attachment(i_poster_file_name, MediaTypeNames.Application.Octet);
            // Add time stamp information for the file.
            ContentDisposition disposition = data.ContentDisposition;
            disposition.CreationDate = System.IO.File.GetCreationTime(i_poster_file_name);
            disposition.ModificationDate = System.IO.File.GetLastWriteTime(i_poster_file_name);
            disposition.ReadDate = System.IO.File.GetLastAccessTime(i_poster_file_name);
            // Add the file attachment to this e-mail message.
            o_mail_poster.Attachments.Add(data);

        }

        /// <summary>Add attachement</summary>
        private void _AddAttachment(string i_attachment_file_name, ref MailMessage o_mail_poster)
        {
            if (i_attachment_file_name == "")
            {
                return;
            }

            if (!File.Exists(i_attachment_file_name))
            {
                return;
            }

            Attachment data = new Attachment(i_attachment_file_name, MediaTypeNames.Application.Octet);
            // Add time stamp information for the file.
            ContentDisposition disposition = data.ContentDisposition;
            disposition.CreationDate = System.IO.File.GetCreationTime(i_attachment_file_name);
            disposition.ModificationDate = System.IO.File.GetLastWriteTime(i_attachment_file_name);
            disposition.ReadDate = System.IO.File.GetLastAccessTime(i_attachment_file_name);
            // Add the file attachment to this e-mail message.
            o_mail_poster.Attachments.Add(data);

        }

        /// <summary>Get list of test Email adresses as a string and as a stack (array). At the moment only one address.
        /// <para>1. Get the Email address from i_newsletter_data</para>
        /// <para>2. Add Email address to o_email_adresses</para>
        /// <para>3. Add Email address to array o_bcc_email_adresses</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Email that shall be sent</param>
        ///  <param name="o_email_adresses">One string with all Email adresses (that will be shown to the user)</param>
        ///  <param name="o_bcc_email_adresses">Array with all Email adresses</param>
        ///  <param name="o_error">Error message</param>
        private bool _GetTestEmailAdresses(NewsletterData i_newsletter_data, out string o_email_adresses, ref Stack<string> o_bcc_email_adresses, out string o_error)
        {
            o_email_adresses = "";
            o_bcc_email_adresses.Clear();
            o_error = "";

            string str_email_address = i_newsletter_data.m_bcc_test;

            if ("" == str_email_address.Trim())
            {
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoBccTest;
                return false;
            }

            o_bcc_email_adresses.Push(str_email_address);

            o_email_adresses = str_email_address;

            return true;
        }

        /// <summary>Get list of recipients for all in the .csv file that has subscribed to the newsletter.
        /// <para>1. Check that the (downloaded) CSV file with addresses exists</para>
        /// <para>2. Create a new Table</para>
        /// <para>3. Fill the Table with the CSV file address data. Call of ToTable.CsvToTable</para>
        /// <para>4. Loop over all addresses in the Table</para>
        /// <para>4.1 Get Field "NewsletterJazz". Call of Table.GetFieldString</para>
        /// <para>4.2 If Field value is "WAHR" get Email address (Field "NewsletterJazz"). Call of Table.GetFieldString</para>
        /// <para>4.3 Add Email address to o_email_adresses separated by comma</para>
        /// <para>4.4 Add Email address to array o_bcc_email_adresses</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Email that shall be sent</param>
        ///  <param name="o_email_adresses">One string with all Email adresses (that will be shown to the user)</param>
        ///  <param name="o_bcc_email_adresses">Array with all Email adresses</param>
        ///  <param name="o_error">Error message</param>
        private bool _GetAllEmailAdresses(NewsletterData i_newsletter_data, out string o_email_adresses, ref Stack<string> o_bcc_email_adresses, out string o_error)
        {
            o_email_adresses = "";
            o_bcc_email_adresses.Clear();
            o_error = "";

            string path_file_name_csv = FileUtil.InputCsvFileName();
            if (!File.Exists(path_file_name_csv))
            {
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoExcelFile + path_file_name_csv;
                return false;
            }

            Table table_addresses = new Table("Jazz addresses");

            if (!ToTable.CsvToTable(path_file_name_csv, ref table_addresses, out o_error))
                return false; ;
 

            // Note start from second row
            for (int i_row = 1; i_row < table_addresses.NumberRows; i_row++) 
            {
                string b_newsletter = table_addresses.GetFieldString(i_row, "NewsletterJazz", out o_error);
                if ("" != o_error)
                {
                    break;
                }

                if (b_newsletter == "WAHR")
                {
                    string str_email_address = table_addresses.GetFieldString(i_row, "e-Mail", out o_error);
                    if ("" != o_error)
                    {
                        break;
                    }

                    if (str_email_address.Trim() != "")
                    {
                        if (i_row == 1)
                            o_email_adresses = o_email_adresses + "\"" + str_email_address + "\"";
                        else if (i_row > 1)
                            o_email_adresses = o_email_adresses + ", \"" + str_email_address + "\"";

                        o_bcc_email_adresses.Push(str_email_address); 
                    }
                }
            } // i_row

            return true;
        }
 
        /// <summary>Get list of recipients as a string and as a stack (array).
        /// <para>1. If i_test_send is true call _GetTestEmailAdresses</para>
        /// <para>2. If i_test_send is false call _GetAllEmailAdresses</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Email that shall be sent</param>
        ///  <param name="i_test_send">Flag telling if only a test Email shall be sent (true) or to all (false)</param>
        ///  <param name="o_email_adresses">One string with all Email adresses (that will be shown to the user)</param>
        ///  <param name="o_bcc_email_adresses">Array with all Email adresses</param>
        ///  <param name="o_error">Error message</param>
        private bool _GetEmailAdresses(NewsletterData i_newsletter_data, bool i_test_send, out string o_email_adresses, ref Stack<string> o_bcc_email_adresses, out string o_error)
        {
            o_email_adresses = "";
            o_bcc_email_adresses.Clear();
            o_error = "";

            if (i_test_send)
            {
                if (!_GetTestEmailAdresses(i_newsletter_data, out o_email_adresses, ref o_bcc_email_adresses, out o_error)) return false;
            }
            else
            {
                if (!_GetAllEmailAdresses(i_newsletter_data, out o_email_adresses, ref o_bcc_email_adresses, out o_error)) return false;
            }

            return true;
        }

        /// <summary> Returns the body HTML text</summary>
        private string _BodyHtml(NewsletterData i_newsletter_data)
        {
            string ret_body_html = "";

            ret_body_html = ret_body_html + i_newsletter_data.Hdr();

            ret_body_html = ret_body_html + i_newsletter_data.ParagraphStart();

            //ret_body_html = ret_body_html + i_newsletter_data.SubjectHtml();
            //ret_body_html = ret_body_html + "<br><br>";

            ret_body_html = ret_body_html + i_newsletter_data.MessageHtml();
            ret_body_html = ret_body_html + "<br>";

            if (i_newsletter_data.GetPosterName() != "" && SendNewsLetter.NewsLetterSettings.Default.PosterEmbedded)
            {
                ret_body_html = ret_body_html + i_newsletter_data.PictureHtml();
                ret_body_html = ret_body_html + "<br>";
            }

            ret_body_html = ret_body_html + i_newsletter_data.InfoHtml();
            ret_body_html = ret_body_html + "<br>";

            if (i_newsletter_data.AddReservationText)
            {
                // ret_body_html = ret_body_html + i_newsletter_data.ReservationHtml();
                ret_body_html = ret_body_html + i_newsletter_data.ReservationInternetHtml();               
                ret_body_html = ret_body_html + "<br>";
            }

            ret_body_html = ret_body_html + i_newsletter_data.EndSubscriptionHtml();
            ret_body_html = ret_body_html + "<br>";

            ret_body_html = ret_body_html + i_newsletter_data.ParagraphEnd();

            return ret_body_html;

        } // _BodyHtml

        /// <summary> Returns the body plain text</summary>
        private string _BodyText(NewsletterData i_newsletter_data)
        {
            string ret_body_txt = "";

            //ret_body_txt = ret_body_txt + "\n\n";
            //ret_body_txt = ret_body_txt + i_newsletter_data.SubjectTxt();
            ret_body_txt = ret_body_txt + "\n\n";

            ret_body_txt = ret_body_txt + i_newsletter_data.MessageTxt();
            ret_body_txt = ret_body_txt + "\n\n";

            ret_body_txt = ret_body_txt + i_newsletter_data.InfoTxt();
            ret_body_txt = ret_body_txt + "\n\n";

            if (i_newsletter_data.AddReservationText)
            {
                // ret_body_txt = ret_body_txt + i_newsletter_data.ReservationTxt(); 
                ret_body_txt = ret_body_txt + i_newsletter_data.ReservationInternetTxt();
                ret_body_txt = ret_body_txt + "\n\n";
            }

            ret_body_txt = ret_body_txt + i_newsletter_data.EndSubscriptionTxt();
            ret_body_txt = ret_body_txt + "\n\n";

            return ret_body_txt;

        } // _BodyTxt

        /// <summary>Constructor creating config file if not existing.</summary>
        public NewsletterMain()
        {
            if (!File.Exists(FileUtil.ConfigFileName()))
            {
                NewsLetterSettings.Default.Save();
            }
        } //NewsletterMain

        /// <summary>Send newsletter to all in the .csv file that has subscribed to the newsletter
        /// <para>1. Send the Emails. Call of _SendNewsLetter with input flags set to all.</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Email that shall be sent</param>
        ///  <param name="o_error">Error message</param>
        public bool SendNewsLetterAll(NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = "";

            bool b_test_send = false;
            bool b_only_to = false;
            if (!_SendNewsLetter(i_newsletter_data, b_test_send, b_only_to, out o_error)) 
                return false;

            return true;
        } // SendNewsLetterAll

        /// <summary>Send test newsletter(s)
        /// <para>1. Send the Emails. Call of _SendNewsLetter with input flags set to test send.</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Email that shall be sent</param>
        ///  <param name="o_error">Error message</param>
        public bool SendNewsLetterTest(NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = "";
            bool b_test_send = true;
            bool b_only_to = false;
            if (!_SendNewsLetter(i_newsletter_data, b_test_send, b_only_to, out o_error))
                return false;

            return true;

        } // SendNewsLetterTest

        /// <summary>Send newsletter only to one person (to To)
        /// <para>1. Send the Email. Call of _SendNewsLetter with input flags set to one person.</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for the Email that shall be sent</param>
        ///  <param name="o_error">Error message</param>
        public bool SendNewsLetterOnePerson(NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = "";
            bool b_test_send = false;
            bool b_only_to = true;
            if (!_SendNewsLetter(i_newsletter_data, b_test_send, b_only_to, out o_error))
                return false;

            return true;
        } // SendNewsLetterAll

    } // NewsletterMain
} // namespace
