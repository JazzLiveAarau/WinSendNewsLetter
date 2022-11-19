using Eml;
using JazzApp;
using JazzFtp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendNewsLetter
{
    /// <summary>Save the newsletter as an EML file and append newsletter data to the XML file JazzNewsletter.xml 
    /// <para>The EML file will be saved in a www/Archive/Newsletter/yyyy directory (where yyyy = ...2018, 2019, 2020, 2021, ... 
    /// Also the attachments will be saved in the same directory.</para>
    /// <para>The file XML file JazzNewsletter is used by the JAZZ live AARAU homepage to display sent newsletters.</para>
    /// <para>For additional information please refer to application NewsletterSave.</para>
    /// <para></para>
    /// </summary>
    public static class NewsletterSave
    {

        #region Global functions and variables

        /// <summary>Server path to the XML newsletter file (JazzNewsletter.xml)</summary>
        private static string m_url_xml_newsletter_file_folder = "XML";
        /// <summary>Get the server path to the XML newsletter file (JazzNewsletter.xml)</summary>
        public static string NewsletterXmlFileServerDir { get { return m_url_xml_newsletter_file_folder; } }

        /// <summary>Get the server path for the backup XML newsletter file (JazzNewsletter_date_time_computer.xml)</summary>
        public static string BackupNewsletterXmlFileServerDir { get { return m_url_archive_newsletter_folder + @"XmlBackup/"; } }

        /// <summary>Name of the XML newsletters file.</summary>
        private static string m_newsletter_xml_filename = "JazzNewsletter.xml";
        /// <summary>Get the name of the XML newsletter file</summary>
        public static string NewsletterXmlFileName { get { return m_newsletter_xml_filename; } }

        /// <summary>Local path for a temporary XML newsletter file (JazzNewsletter.xml)</summary>
        private static string m_local_xml_newsletter_file_folder = "XML";
        /// <summary>Get the local path for a temporary XML newsletter file (JazzNewsletter.xml)</summary>
        private static string NewsletterXmlFileLocalDir { get { return m_local_xml_newsletter_file_folder; } }

        /// <summary>Get the local full name of the temporary XML newsletter file</summary>
        private static string NewsletterXmlLocalFullFileName
        {
            get { return System.Windows.Forms.Application.StartupPath + @"\" + m_local_xml_newsletter_file_folder + @"\" + m_newsletter_xml_filename; }
        }

        /// <summary>Get the local full name of the temporary backup XML newsletter file</summary>
        private static string GetBackupNewsletterXmlLocalFullFileName()
        {
            string ret_file_str = @"";

            string file_name = Path.GetFileNameWithoutExtension(m_newsletter_xml_filename);

            string file_extension = Path.GetExtension(m_newsletter_xml_filename);

            string file_backup = file_name + @"_" + TimeUtil.YearMonthDayHourMinSecComputer() + file_extension;

            ret_file_str = ret_file_str + System.Windows.Forms.Application.StartupPath + @"\";

            ret_file_str = ret_file_str + m_local_xml_newsletter_file_folder + @"\";

            ret_file_str = ret_file_str + file_backup;

            return ret_file_str;

        }  // GetBackupNewsletterXmlLocalFullFileName

        /// <summary>Start path to the server directory where newsletter files are stored</summary>
        private static string m_url_archive_newsletter_folder = "Archive/Newsletter/";
        /// <summary>Get the start path to the server directory where newsletter files are stored</summary>
        public static string ArchiveNewsletterServerDir { get { return m_url_archive_newsletter_folder; } }

        /// <summary>Returns the path to the server directory where newsletter files are stored</summary>
        public static string GetArchiveNewsletterYearServerDir(int i_send_year)
        {
            return ArchiveNewsletterServerDir + i_send_year.ToString() + @"/";

        } // GetArchiveNewsletterYearServerDir

        /// <summary>Full local path for a temporary EML newsletter file</summary>
        private static string m_local_eml_newsletter_file_folder = System.Windows.Forms.Application.StartupPath + @"\EML\";
        /// <summary>Get the full local path for a temporary EML newsletter file</summary>
        private static string NewsletterEmlFileLocalDir { get { return m_local_eml_newsletter_file_folder; } }

        /// <summary>Returns the server ftp password</summary>
        private static string FtpPassword = @"42SN4bX9";

        /// <summary>Returns the path to the local execution directory</summary>
        private static string ExeDirectory = System.Windows.Forms.Application.StartupPath;

        #endregion // Global functions and variables

        #region Main execution function

        /// <summary>Save newsletter as EML and append newsletter data to the XML file JazzNewsletter.xml 
        /// <para>1. Create JazzEml object and get data from input NewsletterData object. Call of GetNewsletterData</para>
        /// <para>2. Create and save EML file. Call of SaveEml.Execute</para>
        /// <para>3. Append newsletter data from the JazzEml object to the XML object and create local JazzNewsletter.xml file.
        /// Call of AppendEmlToXmlSaveLocally</para>
        /// <para>4. Upload XML, EML, image and attachment files to the server. Call of UploadXmlEmlAttachmentFiles</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for an Email that have been sent</param>
        ///  <param name="o_error">Error message</param>
        public static bool Execute(NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = @"";

            string error_msg = @"";

            JazzEml sent_eml = GetNewsletterData(i_newsletter_data, out error_msg);
            if (sent_eml == null || error_msg.Length > 0)
            {
                o_error = @"NewsletterSave.Execute GetNewsletterData failed " + error_msg;

                return false;
            }

            if (!SaveEml.Execute(ref sent_eml, out error_msg))
            {
                o_error = @"NewsletterSave.Execute SaveEml.Execute failed " + error_msg;

                return false;
            }

            if (!AppendEmlToXmlSaveLocally(sent_eml, out error_msg))
            {
                o_error = @"NewsletterSave.Execute AppendEmlToXmlSaveLocally failed " + error_msg;

                return false;
            }

            if (!UploadXmlEmlAttachmentFiles(sent_eml, out error_msg))
            {
                o_error = @"NewsletterSave.Execute UploadXmlEmlAttachmentFiles failed " + error_msg;

                return false;
            }

            return true;

        } // Execute

        #endregion // Main execution function

        #region Set a JazzEml object

        /// <summary>Returns an JazzEml object with data from a NewsletterData object
        /// <para>1. Set the send date. Call of GetNewsletterDataSendDate</para>
        /// <para>2. Set the header date. Call of GetNewsletterDataHeader</para>
        /// <para>3. Set the embedded image data. Call of GetNewsletterDataEmbeddedImage</para>
        /// <para>4. Set the attached file data. Call of GetNewsletterDataAttachedFile</para>
        /// <para>5. Set the message text for the newsletter. Call of GetNewsletterDataMessage</para>
        /// <para>6. Set the EML file data. Call of SetEmlFileData</para>
        /// </summary>
        ///  <param name="i_newsletter_data">Holds data for an Email that have been sent</param>
        ///  <param name="o_error">Error message</param>
        ///  <returns>JazzEml object. For error null is returned and/or o_error is set</returns>
        private static JazzEml GetNewsletterData(NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = @"";

            JazzEml ret_eml = null;

            string error_msg = @"";

            if (i_newsletter_data == null)
            {
                o_error = @"NewsletterSave.GetNewsletterData Input NewsletterData object is null";

                return ret_eml;
            }

            ret_eml = new JazzEml();

            if (!GetNewsletterDataSendDate(ref ret_eml, i_newsletter_data, out error_msg))
            {
                o_error = @"NewsletterSave.GetNewsletterData " + error_msg;

                return ret_eml;
            }

            if (!GetNewsletterDataHeader(ref ret_eml, i_newsletter_data, out error_msg))
            {
                o_error = @"NewsletterSave.GetNewsletterData " + error_msg;

                return ret_eml;
            }

            if (!GetNewsletterDataEmbeddedImage(ref ret_eml, i_newsletter_data, out error_msg))
            {
                o_error = @"NewsletterSave.GetNewsletterData " + error_msg;

                return ret_eml;
            }

            if (!GetNewsletterDataAttachedFile(ref ret_eml, i_newsletter_data, out error_msg))
            {
                o_error = @"NewsletterSave.GetNewsletterData " + error_msg;

                return ret_eml;
            }

            if (!GetNewsletterDataMessage(ref ret_eml, i_newsletter_data, out error_msg))
            {
                o_error = @"NewsletterSave.GetNewsletterData " + error_msg;

                return ret_eml;
            }

            if (!SetEmlFileData(ref ret_eml, out error_msg))
            {
                o_error = @"NewsletterSave.GetNewsletterData " + error_msg;

                return ret_eml;
            }

            return ret_eml;

        } // GetNewsletterData

        /// <summary>Sets the send date JazzEml variables: 
        /// <para>SendYear, SendMonth and SendDay</para>
        /// <para>Function TimeUtil.YearMonthDayArray is called</para>
        /// </summary>
        /// <param name="ref_eml">JazzEml object with the variables that will be set</param>
        /// <param name="i_newsletter_data">Holds data for an Email that have been sent</param>
        /// <param name="o_error">Error message</param>
        private static bool GetNewsletterDataSendDate(ref JazzEml ref_eml, NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = @"";

            int[] current_date_array = TimeUtil.YearMonthDayArray();

            ref_eml.SendYear = current_date_array[0];

            ref_eml.SendMonth = current_date_array[1];

            ref_eml.SendDay = current_date_array[2];

            return true;

        } // GetNewsletterDataSendDate

        /// <summary>Sets the header data JazzEml variables: 
        /// <para>From, To and Subject</para>
        /// </summary>
        /// <param name="ref_eml">JazzEml object with the variables that will be set</param>
        /// <param name="i_newsletter_data">Holds data for an Email that have been sent</param>
        /// <param name="o_error">Error message</param>
        private static bool GetNewsletterDataHeader(ref JazzEml ref_eml, NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = @"";

            ref_eml.From = i_newsletter_data.m_from;

            ref_eml.To = i_newsletter_data.m_to;

            ref_eml.Subject = i_newsletter_data.SubjectHtml();

            return true;

        } // GetNewsletterDataHeader

        /// <summary>Sets the embedded image JazzEml variables: 
        /// <para>AttachmentImageLocal, AttachmentImageServer, AttachmentImagePathServer and EmbeddedPosterFlag</para>
        /// </summary>
        /// <param name="ref_eml">JazzEml object with the variables that will be set</param>
        /// <param name="i_newsletter_data">Holds data for an Email that have been sent</param>
        /// <param name="o_error">Error message</param>
        private static bool GetNewsletterDataEmbeddedImage(ref JazzEml ref_eml, NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = @"";

            ref_eml.AttachmentImageLocal = i_newsletter_data.GetPosterName();

            if (ref_eml.AttachmentImageLocal.Length > 0)
            {
                string image_file_name = Path.GetFileName(ref_eml.AttachmentImageLocal);

                ref_eml.AttachmentImageServer = image_file_name;

                if (ref_eml.SendYear < 1995)
                {
                    o_error = @"GetNewsletterDataEmbeddedImage JazzEml.SendYear is not set";

                    return false;
                }

                ref_eml.AttachmentImagePathServer = GetArchiveNewsletterYearServerDir(ref_eml.SendYear);

                ref_eml.EmbeddedPosterFlag = true;
            }

            return true;

        } // GetNewsletterDataEmbeddedImage

        /// <summary>Sets the attached file JazzEml variables: 
        /// <para>AttachmentFileLocal, AttachmentFileServer and AttachmentPathServer </para>
        /// </summary>
        /// <param name="ref_eml">JazzEml object with the variables that will be set</param>
        /// <param name="i_newsletter_data">Holds data for an Email that have been sent</param>
        /// <param name="o_error">Error message</param>
        private static bool GetNewsletterDataAttachedFile(ref JazzEml ref_eml, NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = @"";

            ref_eml.AttachmentFileLocal = i_newsletter_data.GetAttachmentName();

            if (ref_eml.AttachmentFileLocal.Length > 0)
            {
                string attachment_file_name = Path.GetFileName(ref_eml.AttachmentFileLocal);

                ref_eml.AttachmentFileServer = attachment_file_name;

                if (ref_eml.SendYear < 1995)
                {
                    o_error = @"GetNewsletterDataAttachedFile JazzEml.SendYear is not set";

                    return false;
                }

                ref_eml.AttachmentPathServer = GetArchiveNewsletterYearServerDir(ref_eml.SendYear);
            }

            return true;

        } // GetNewsletterDataAttachedFile

        /// <summary>Sets the EML file variables of the object JazzEml: 
        /// <para>EmlFileNameLocal, EmlFileServer and EmlFileServer</para>
        /// <para>Function GetArchiveNewsletterYearServerDir is called</para>
        /// </summary>
        /// <param name="ref_eml">JazzEml object with the variables that will be set</param>
        /// <param name="o_error">Error message</param>
        private static bool SetEmlFileData(ref JazzEml ref_eml, out string o_error)
        {
            o_error = @"";

            string eml_local_dir = NewsletterEmlFileLocalDir;

            if (!Directory.Exists(eml_local_dir))
            {
                Directory.CreateDirectory(eml_local_dir);
            }

            string eml_file_name = GetEmlFileName(ref_eml);

            string eml_full_file_name = eml_local_dir + eml_file_name;

            string eml_server_dir = GetArchiveNewsletterYearServerDir(ref_eml.SendYear);

            ref_eml.EmlFileNameLocal = eml_full_file_name;

            ref_eml.EmlFileServer = eml_file_name;

            ref_eml.EmlPathServer = eml_server_dir;

            return true;

        } // SetEmlFileData

        /// <summary>Returns EML file name dYYYYMMDD.eml where YYYYMMDD is the send date</summary>
        private static string GetEmlFileName(JazzEml i_eml)
        {
            string ret_file_name = @"";

            string send_year = i_eml.SendYear.ToString();

            string send_month = i_eml.SendMonth.ToString();
            if (send_month.Length == 1)
            {
                send_month = @"0" + send_month;
            }

            string send_day = i_eml.SendDay.ToString();
            if (send_day.Length == 1)
            {
                send_day = @"0" + send_day;
            }

            ret_file_name = ret_file_name + @"d";

            ret_file_name = ret_file_name + send_year;

            ret_file_name = ret_file_name + send_month;

            ret_file_name = ret_file_name + send_day;

            ret_file_name = ret_file_name + @".eml";

            return ret_file_name;

        } // GetEmlFileName

        /// <summary>Sets the header data JazzEml variable MsgHtml
        /// <para>1. Get the text for the newsletter from the input NewsletterData object</para>
        /// <para>2. Construct the embedded picture HTML string if an image is attached.</para>
        /// <para>3. Construct the reservation HTML string. 
        /// Call of NewsletterData.AddReservationText and NewsletterData.ReservationInternetHtml</para>
        /// <para>4. Get the homepage link HTML string. Call of NewsletterData.InfoHtml</para>
        /// <para>5. Get the end-subscribe HTML string. Call of NewsletterData.EndSubscriptionHtml</para>
        /// <para>6. Combine all HTML strings and set JazzEml.MsgHtml</para>
        /// <para></para>
        /// </summary>
        /// <param name="ref_eml">JazzEml object with the variables that will be set</param>
        /// <param name="i_newsletter_data">Holds data for an Email that have been sent</param>
        /// <param name="o_error">Error message</param>
        private static bool GetNewsletterDataMessage(ref JazzEml ref_eml, NewsletterData i_newsletter_data, out string o_error)
        {
            o_error = @"";

            string msg_html = i_newsletter_data.MessageHtml();

            string img_html = "";

            if (ref_eml.AttachmentImageLocal.Length > 0)
            {
                img_html = ref_eml.GetImgHtmlString();

                if (img_html.Length == 0)
                {
                    o_error = @"GetNewsletterDataMessage JazzEml.GetImgHtmlString failed";

                    return false;
                }

                img_html = img_html + "<br>";
            }

            string reservation_html = @"";

            if (i_newsletter_data.AddReservationText)
            {
                reservation_html = reservation_html + i_newsletter_data.ReservationInternetHtml();
            }

            string info_html = i_newsletter_data.InfoHtml();

            string end_subscribe_html = i_newsletter_data.EndSubscriptionHtml();

            ref_eml.MsgHtml = msg_html + img_html + reservation_html + info_html + end_subscribe_html;

            return true;

        } // GetNewsletterDataMessage

        #endregion // Set a JazzEml object

        #region Append the JazzEml object to XML

        /// <summary>Appends a JazzEml object to the newsletter XML object and save the updated XML file in local directory
        /// <para>1. Get the XML object for the JazzNewsletter.xml file on the server. Call of InitXmlNewsletter</para>
        /// <para>2. Create a backup of the JazzNewsletter.xml file on the server. Call of CreateBackupNewsletterXmlFile</para>
        /// <para>3 Create (instantiate) a JazzNewsletter object</para>
        /// <para>4 Set variables of the JazzNewsletter from the JazzEml object </para>
        /// <para>5 Set undefined values to NotYetSetNodeValue. Call of JazzNewsletter.GetObjectWithUndefinedNodeValues</para>
        /// <para>6 Check variable values. Call of JazzNewsletter.CheckParameterValues</para>
        /// <para>7 Add the JazzNewsletter object to the JazzNewsletter.xml XML object. Call of JazzXml.AddNewsletter</para>
        /// <para>8. Create a local JazzNewsletter.xml file in subdirectory XML. Call of JazzXml.WriteToFile</para>
        /// <para></para>
        /// </summary>
        /// <param name="i_eml_object">JazzEml objects</param>
        /// <param name="o_error">Error message</param>
        static private bool AppendEmlToXmlSaveLocally(JazzEml i_eml_object, out string o_error)
        {
            o_error = @"";

            string error_msg = @"";

            if (!InitXmlNewsletter(out error_msg))
            {
                o_error = @"NewsletterSave.AppendEmlToXmlSaveLocally InitXmlNewsletter failed " + error_msg;

                return false;
            }

            if (!CreateBackupNewsletterXmlFile(out error_msg))
            {
                o_error = @"NewsletterSave.AppendEmlToXmlSaveLocally CreateBackupNewsletterXmlFile failed " + error_msg;

                return false;
            }

            JazzNewsletter jazz_newsletter = new JazzNewsletter();

            jazz_newsletter.EmlPathServer = i_eml_object.EmlPathServer;

            jazz_newsletter.EmlFileServer = i_eml_object.EmlFileServer;

            jazz_newsletter.SendYearInt = i_eml_object.SendYear;

            jazz_newsletter.SendMonthInt = i_eml_object.SendMonth;

            jazz_newsletter.SendDayInt = i_eml_object.SendDay;

            jazz_newsletter.Subject = i_eml_object.Subject;

            jazz_newsletter.From = i_eml_object.From;

            jazz_newsletter.To = i_eml_object.To;

            jazz_newsletter.MsgHtml = i_eml_object.MsgHtml;

            jazz_newsletter.AttachmentImagePathServer = i_eml_object.AttachmentImagePathServer;

            jazz_newsletter.AttachmentImageServer = i_eml_object.AttachmentImageServer;

            jazz_newsletter.EmbeddedPosterFlagBoolean = i_eml_object.EmbeddedPosterFlag;

            jazz_newsletter.AttachmentPathServer = i_eml_object.AttachmentPathServer;

            jazz_newsletter.AttachmentFileServer = i_eml_object.AttachmentFileServer;


            jazz_newsletter = jazz_newsletter.GetObjectWithUndefinedNodeValues(jazz_newsletter);

            if (!jazz_newsletter.CheckParameterValues(out error_msg))
            {
                o_error = "NewsletterSave.AppendEmlToXmlSaveLocally JazzNewsletter.CheckParameterValues failed " + error_msg;

                return false;
            }

            if (!JazzXml.AddNewsletter(jazz_newsletter, out error_msg))
            {
                o_error = "NewsletterSave.AppendEmlToXmlSaveLocally JazzXml.AddNewsletter failed " + error_msg;

                return false;
            }

            string full_file_name = NewsletterXmlLocalFullFileName;

            if (!JazzXml.WriteToFile(JazzXml.GetObjectNewsletter(), full_file_name, out error_msg))
            {
                o_error = @"NewsletterSave.AppendEmlToXmlSaveLocally  JazzXml.WriteToFile failed" + error_msg;

                return false;
            }

            return true;

        } // AppendEmlToXmlSaveLocally

        /// <summary>Initialization of newsletter object corresponding to the XML file JazzNewsletter.xml
        /// <para>1. Set FTP connection data. Call of JazzXml.SetFtpConnectionData</para>
        /// <para>2. Create the XML for the current JazzNewsletter.xml file on the server.
        /// Call of JazzXml.InitNewsletter. This function also creates the local directory XML</para>
        /// <para></para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        static private bool InitXmlNewsletter(out string o_error)
        {
            o_error = @"";

            bool ret_init = true;

            JazzXml.SetFtpConnectionData(NewsLetterSettings.Default.FtpHost, NewsLetterSettings.Default.FtpUser, FtpPassword, ExeDirectory);

            string error_message = @"";

            if (!JazzXml.InitNewsletter(NewsletterXmlFileServerDir, NewsletterXmlFileName, out error_message))
            {
                o_error = @"NewsletterSave.InitXmlNewsletter JazzXml.InitNewsletter failed " + error_message;

                return false;
            }

            return ret_init;

        } // InitXmlNewsletter

        /// <summary>Create and upload a backup JazzNewsletter.xml file
        /// <para>1. Get the local backup file name. Call of GetBackupNewsletterXmlLocalFullFileName</para>
        /// <para>2. Create the local backup file. Call of JazzXml.WriteToFile</para>
        /// <para>3. Upload the file to the server. Call of BackupNewsletterXmlFileServerDir and UploadOneFileToServer</para>
        /// <para>4. Delete the local backup file. Call of File.Exists and File.Delete</para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        static private bool CreateBackupNewsletterXmlFile(out string o_error)
        {
            o_error = @"";

            string error_msg = @"";

            string full_file_name = GetBackupNewsletterXmlLocalFullFileName();

            if (!JazzXml.WriteToFile(JazzXml.GetObjectNewsletter(), full_file_name, out error_msg))
            {
                o_error = @"NewsletterSave.CreateBackupNewsletterXmlFile  JazzXml.WriteToFile failed" + error_msg;

                return false;
            }

            string file_name = Path.GetFileName(full_file_name);

            string backup_server_dir = BackupNewsletterXmlFileServerDir;

            if (!UploadOneFileToServer(file_name, backup_server_dir, full_file_name, out error_msg))
            {
                o_error = @"NewsletterSave.CreateBackupNewsletterXmlFile UploadOneFileToServer failed (backup file) " + error_msg;

                return false;
            }

            if (File.Exists(full_file_name))
            {
                File.Delete(full_file_name);
            }

            return true;

        }  // CreateBackupNewsletterXmlFile

        #endregion // Append the JazzEml object to XML

        #region Upload XML and all other files to the server

        /// <summary>Upload XML, EML, image and attachment files to the server
        /// <para>1. Upload the image file if existing. Call of UploadOneFileToServer</para>
        /// <para>2. Upload the attachment file if existing. Call of UploadOneFileToServer</para>
        /// <para>3. Upload the EML file if existing. Call of UploadOneFileToServer</para>
        /// <para>4. Upload the XML file JazzNewsletter.xml. Call of UploadOneFileToServer</para>
        /// <para>5. Delete the local XML file JazzNewsletter.xml. File.Exists and File.Delete</para>
        /// <para>6. Delete the local EML file. File.Exists and File.Delete</para>
        /// </summary>
        ///  <param name="i_eml_object">Holds data for an Email that have been sent</param>
        ///  <param name="o_error">Error message</param>
        static private bool UploadXmlEmlAttachmentFiles(JazzEml i_eml_object, out string o_error)
        {
            o_error = @"";

            string error_msg = @"";

            if (i_eml_object.AttachmentImageServer.Length > 4)
            {
                if (!UploadOneFileToServer(i_eml_object.AttachmentImageServer, i_eml_object.AttachmentImagePathServer, i_eml_object.AttachmentImageLocal, out error_msg))
                {
                    o_error = @"NewsletterSave.UploadXmlEmlAttachmentFiles UploadOneFileToServer failed (attached image) " + error_msg;

                    return false;
                }

            } // Embedded image file exists

            if (i_eml_object.AttachmentFileServer.Length > 4)
            {
                if (!UploadOneFileToServer(i_eml_object.AttachmentFileServer, i_eml_object.AttachmentPathServer, i_eml_object.AttachmentFileLocal, out error_msg))
                {
                    o_error = @"NewsletterSave.UploadXmlEmlAttachmentFiles UploadOneFileToServer failed (attached file) " + error_msg;

                    return false;
                }

            } // Attached file exists

            if (i_eml_object.EmlFileServer.Length > 4)
            {
                if (!UploadOneFileToServer(i_eml_object.EmlFileServer, i_eml_object.EmlPathServer, i_eml_object.EmlFileNameLocal, out error_msg))
                {
                    o_error = @"NewsletterSave.UploadXmlEmlAttachmentFiles UploadOneFileToServer failed (EML file) " + error_msg;

                    return false;
                }

            } // EML file exists

            string full_file_name = NewsletterXmlLocalFullFileName;

            string file_name = NewsletterXmlFileName;

            string server_dir = NewsletterXmlFileServerDir;

            if (!UploadOneFileToServer(file_name, server_dir, full_file_name, out error_msg))
            {
                o_error = @"NewsletterSave.UploadXmlEmlAttachmentFiles UploadOneFileToServer failed (XML file) " + error_msg;

                return false;
            }

            if (File.Exists(full_file_name))
            {
                File.Delete(full_file_name);
            }

            if (File.Exists(i_eml_object.EmlFileNameLocal))
            {
                File.Delete(i_eml_object.EmlFileNameLocal);
            }

            return true;

        } // UploadXmlEmlAttachmentFiles

        /// <summary>Upload one file to the server
        /// <para>1. Extract path and directory from the input full file name. Remove execution path since it is set in JazzFtp.Input
        /// Call of Path.GetFileName, Path.GetDirectoryName, string.IndexOf and string.Substring</para>
        /// <para>2. Modify input server path since this is for the homepage. Call of ModifyServerDirectoryName</para>
        /// <para>3. Create server directory (/year/) if not existing. Call of CreateServerDirIfNotExisting</para>
        /// <para>4. Upload the file. Create JazzFtp.Input (case UpLoadFile) object and call of JazzFtp.Execute.Run</para>
        /// </summary>
        ///  <param name="i_file_name_server">Server name of the file that shall be uploaded</param>
        ///  <param name="i_server_dir">Server directory name for the file that shall be uploaded</param>
        ///  <param name="i_full_file_name_local">Full local file name for the file that shall be uploaded</param>
        ///  <param name="o_error">Error message</param>
        static private bool UploadOneFileToServer(string i_file_name_server, string i_server_dir, string i_full_file_name_local, out string o_error)
        {
            o_error = @"";

            string error_msg = @"";

            string file_name_local = Path.GetFileName(i_full_file_name_local);

            string full_local_dir = Path.GetDirectoryName(i_full_file_name_local);

            int index_exe_dir = full_local_dir.IndexOf(ExeDirectory);

            if (index_exe_dir != 0)
            {
                o_error = @"NewsletterSave.UploadOneFileToServer ExeDirectory not start of " + i_full_file_name_local;

                return false;
            }

            string local_dir = full_local_dir.Substring(ExeDirectory.Length + 1);

            string server_dir = ModifyServerDirectoryName(i_server_dir);

            if (!CreateServerDirIfNotExisting(server_dir, out error_msg))
            {
                o_error = @"NewsletterSave.UploadOneFileToServer CreateServerDirIfNotExisting failed " + error_msg;

                return false;
            }

            JazzFtp.Input ftp_input_upload = new JazzFtp.Input(ExeDirectory, Input.Case.UpLoadFile);

            ftp_input_upload.ServerDirectory = server_dir;
            ftp_input_upload.ServerFileName = i_file_name_server;

            ftp_input_upload.LocalDirectory = local_dir;
            ftp_input_upload.LocalFileName = file_name_local;

            JazzFtp.Result ftp_result_upload = JazzFtp.Execute.Run(ftp_input_upload);

            if (!ftp_result_upload.Status)
            {
                o_error = @"NewsletterSave.UploadOneFileToServer JazzFtp.Execute.Run failed " + ftp_result_upload.ErrorMsg;

                return false;
            }

            return true;

        } // UploadOneFileToServer

        /// <summary>Modify the server directory name
        /// <para>In JazzEml there is one en slash that will be removed.</para>
        /// <para>The relative server directory URL in JazzEml is for the homepage. www is added</para>
        /// </summary>
        ///  <param name="i_server_dir">Server directory</param>
        static private string ModifyServerDirectoryName(string i_server_dir)
        {
            string ret_dir_str = i_server_dir;

            string last_char_server_dir = i_server_dir.Substring(i_server_dir.Length - 1);

            if (last_char_server_dir.Equals("/"))
            {
                ret_dir_str = i_server_dir.Substring(0, i_server_dir.Length - 1);
            }

            ret_dir_str = "www/" + ret_dir_str;

            return ret_dir_str;

        } // ModifyServerDirectoryName

        /// <summary>Create server directory if not existing
        /// <para>1. </para>
        /// </summary>
        ///  <param name="i_server_dir">Server directory</param>
        ///  <param name="o_error">Error message</param>
        static private bool CreateServerDirIfNotExisting(string i_server_dir, out string o_error)
        {
            o_error = @"";

            string error_msg = @"";

            bool dir_exists = false;

            if (!CheckIfServerDirExist(i_server_dir, out dir_exists, out error_msg))
            {
                o_error = @"NewsletterSave.CreateServerDirIfNotExisting CheckIfServerDirExist failed " + error_msg;

                return false;
            }

            if (dir_exists)
            {
                // Directory already exists
                return true;
            }

            JazzFtp.Input ftp_input_dir_create = new JazzFtp.Input(ExeDirectory, JazzFtp.Input.Case.DirCreate);

            ftp_input_dir_create.ServerDirectory = i_server_dir;

            JazzFtp.Result result_create = JazzFtp.Execute.Run(ftp_input_dir_create);
            if (!result_create.Status)
            {
                o_error = @"NewsletterSave.UploadOneFileToServer JazzFtp.Execute.Run (DirCreate) failed ";

                return false;
            }

            return true;

        } // CreateServerDirIfNotExisting

        /// <summary>Check if server dir exists
        /// <para>1. </para>
        /// </summary>
        ///  <param name="i_server_dir">Server directory</param>
        ///  <param name="o_dir_exists">Flag telling if the directory exists</param>
        ///  <param name="o_error">Error message</param>
        static private bool CheckIfServerDirExist(string i_server_dir, out bool o_dir_exists, out string o_error)
        {
            o_dir_exists = false;

            o_error = @"";

            JazzFtp.Input ftp_input_dir_exist = new JazzFtp.Input(ExeDirectory, JazzFtp.Input.Case.DirExists);

            ftp_input_dir_exist.ServerDirectory = i_server_dir;

            JazzFtp.Result result_exists = JazzFtp.Execute.Run(ftp_input_dir_exist);
            if (!result_exists.Status)
            {
                o_error = @"NewsletterSave.UploadOneFileToServer JazzFtp.Execute.Run (DirExists) failed ";

                return false;
            }

            if (!result_exists.BoolResult)
            {
                o_dir_exists = false;

                return true;
            }
            else
            {
                o_dir_exists = true;
            }

            return true;

        } // CheckIfServerDirExist

        #endregion // Upload XML and all other files to the server

    } // NewsletterSave

} // namespace
