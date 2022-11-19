using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Resources;
using System.IO;
using JazzVersion;

using System.Collections.Generic;

namespace SendNewsLetter
{
	/// <summary>Main form for the SendNewsLetter application</summary>
	public class NewsletterForm : Form
    {
        #region Member variables controls

        private TextBox m_textbox_to;
		private Button m_button_all_send;
		private Label m_label_to;
		private Label m_label_subject;
		private TextBox m_textbox_subject;
        private RichTextBox m_rich_textbox_body;
		private MainMenu mMainMenu;
		private ContextMenu cmExtras;
		private MenuItem cmiPaste;
        private MenuItem cmiClearAll;
        private ImageList ImageList;
        private Label m_label_from;
		private Label m_label_poster;
		private PictureBox m_picture_logo;
		private Button m_button_ende;
		private System.ComponentModel.IContainer components;

        private TextBox m_textbox_message;
        private Label m_label_test_address;
        private Button m_button_test_send;
        private Label labelText;
        private DateTimePicker m_date_time_picker;
        private ComboBox m_combo_box_hour;
        private ComboBox m_combo_box_minute;
        private Label m_label_datum;
        private Label m_label_zeit;
        private Label m_label_colone;
        private TextBox m_text_box_band;
        private Label m_label_band;
        private ComboBox m_combo_box_from;
        private Button m_button_subject_combine;
        private ComboBox m_combo_box_pics;
        private ComboBox m_combo_box_bcc_test;
        private PictureBox m_picture_box;
        private RichTextBox m_rich_text_box_black;

        #endregion

        /// <summary> Main class for the SendNewsLetter application</summary>
        private NewsletterMain m_newsletter_main;

        /// <summary>Picture file names</summary>
        private string[] m_picture_file_names;

        /// <summary>Downloaded picture file names</summary>
        private string[] m_downloaded_picture_file_names;

        /// <summary>Attachment file names</summary>
        private string[] m_attachment_file_names;

        /// <summary>Bitmap for the poster</summary>
        private Bitmap m_bitmap_poster;

        /// <summary>Debug output flag</summary>
        private bool m_debug_output = false;
        private ComboBox m_combo_box_attachment;
        private Label m_label_attachment;
        private Button m_button_help;
        private Button m_button_edit_config;

        #region Create ToolTips
        private ToolTip m_tool_tip_application = new ToolTip();
        private ToolTip m_tool_tip_reset = new ToolTip();
        private ToolTip m_tool_tip_band = new ToolTip();
        private ToolTip m_tool_tip_date = new ToolTip();
        private ToolTip m_tool_tip_hour = new ToolTip();
        private ToolTip m_tool_tip_minute = new ToolTip();
        private ToolTip m_tool_tip_pic = new ToolTip();
        private ToolTip m_tool_tip_subject = new ToolTip();
        private ToolTip m_tool_tip_picture_box = new ToolTip();
        private ToolTip m_tool_tip_text = new ToolTip();
        private ToolTip m_tool_tip_only_to = new ToolTip();
        private ToolTip m_tool_tip_test_send = new ToolTip();
        private ToolTip m_tool_tip_send_all = new ToolTip();
        private ToolTip m_tool_tip_text_box_to = new ToolTip();
        private ToolTip m_tool_tip_combo_box_from = new ToolTip();
        private ToolTip m_tool_tip_combo_box_test = new ToolTip();
        private ToolTip m_tool_tip_end_session = new ToolTip();
        private ToolTip m_tool_tip_attachment = new ToolTip();
        private ToolTip m_tool_tip_edit_config = new ToolTip();
        private ToolTip m_tool_tip_help = new ToolTip();
        private ToolTip m_tool_tip_update = new ToolTip();
        private ToolTip m_tool_tip_only_downloaded_pics = new ToolTip();
        private ToolTip m_tool_tip_checkbox_with_reservation_text = new ToolTip();
        #endregion

        private Button m_button_update;
        private CheckBox m_checkbox_only_down_loaded_pics;
        private Label m_label_version;
        private CheckBox m_checkbox_with_reservation_text;

        /// <summary>Flag telling if the user manually sets the subject</summary>
        private bool m_update_subject = true;

        /// <summary>Constructor: Initialization of controls, creates NewsletterMain</summary>
		public NewsletterForm()
		{
			//CultureInfo ci =new CultureInfo("en-us");
            CultureInfo ci = new CultureInfo("de-CH");
			Thread.CurrentThread.CurrentCulture = ci;
			Thread.CurrentThread.CurrentUICulture = ci;

            //System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("de-CH");
            //System.Threading.Thread.CurrentThread.CurrentUICulture = System.Threading.Thread.CurrentThread.CurrentCulture;

			InitializeComponent();

            m_newsletter_main = new NewsletterMain();

            m_debug_output = NewsLetterSettings.Default.DebugOutput;

            string error_message = "";
            bool b_posters = false;
            bool b_new_version_templates = false;
            if (!m_newsletter_main.DownloadFiles(b_new_version_templates, out error_message))
            {
                MessageBox.Show(error_message);
            }
            else
            {
                b_posters = true;
            }

            if (!m_newsletter_main.InitXml(out error_message))
            {
                MessageBox.Show(error_message);
            }

            string error_message_names = "";
            bool b_names = m_newsletter_main.GetDownloadedPictureNames(out m_downloaded_picture_file_names, out error_message_names);

            _InitializeControls();

            _VersionCheck();

            _SetToolTips();

            FileUtil.HelpFileCreateIfMissing();

            bool b_addresses = false;
            if (!m_newsletter_main.DownloadAddresseFile(out error_message))
            {
                MessageBox.Show(error_message);
            }
            else
            {
                b_addresses = true;
            }

            if (!b_addresses)
            {
                this.m_textbox_message.Text = NewsLetterSettings.Default.ErrMsgNoExcelFileDownload;
            }
            else if (b_addresses && !b_posters)
            {
                this.m_textbox_message.Text = NewsLetterSettings.Default.MsgExcelFileDownload;
            }
            else if (b_addresses && b_posters)
            {
                this.m_textbox_message.Text = NewsLetterSettings.Default.MsgExcelPosterFilesDownload;
            }

            NewsLetterXml.SetBandToNextConcert(m_text_box_band);

        } // Constructor

        /// <summary>Checks if there is anew version is available
        /// <para>Message will be written to the version label if available.</para>
        /// </summary>
        private void _VersionCheck()
        {
            VersionInput version_input = new VersionInput();

            version_input.ExeDirectory = System.Windows.Forms.Application.StartupPath;
            version_input.VersionString = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

            string error_message = @"";
            if (!VersionUtil.Init(version_input, out error_message))
            {
                m_label_version.Text = error_message;
                return;
            }

            bool new_version = false;
            string version_str = @"";
           
            if (!VersionUtil.NewVersionIsAvailable(out new_version, out version_str, out error_message))
            {
                this.m_label_version.Text = error_message;
                return;
            }

            if (new_version)
            {
                this.m_label_version.Text = NewsLetterSettings.Default.MsgNewVersionIsAvailable + version_str;
            }
            else
            {
                this.m_label_version.Text = NewsLetterSettings.Default.GuiLabelVersion + version_str;
            }

        } // _VersionCheck

        /// <summary>Set tool tips</summary>
        private void _SetToolTips()
        {
            m_tool_tip_application.SetToolTip(this, NewsLetterSettings.Default.ToolTipApplication);
            ToolTipUtil.SetDelays(ref m_tool_tip_application);
            m_tool_tip_reset.SetToolTip(this.m_button_subject_combine, NewsLetterSettings.Default.ToolTipButtonCombineSubject);
            ToolTipUtil.SetDelays(ref m_tool_tip_reset);
            m_tool_tip_band.SetToolTip(this.m_text_box_band, NewsLetterSettings.Default.ToolTipTextBoxBand);
            ToolTipUtil.SetDelays(ref m_tool_tip_band);
            m_tool_tip_date.SetToolTip(this.m_date_time_picker, NewsLetterSettings.Default.ToolTipDateTimePicker);
            ToolTipUtil.SetDelays(ref m_tool_tip_date);
            m_tool_tip_hour.SetToolTip(this.m_combo_box_hour, NewsLetterSettings.Default.ToolTipComboBoxHour);
            ToolTipUtil.SetDelays(ref m_tool_tip_hour);
            m_tool_tip_minute.SetToolTip(this.m_combo_box_minute, NewsLetterSettings.Default.ToolTipComboBoxMinute);
            ToolTipUtil.SetDelays(ref m_tool_tip_minute);
            m_tool_tip_pic.SetToolTip(this.m_combo_box_pics, NewsLetterSettings.Default.ToolTipComboBoxPics);
            ToolTipUtil.SetDelays(ref m_tool_tip_pic);
            m_tool_tip_subject.SetToolTip(this.m_textbox_subject, NewsLetterSettings.Default.ToolTipTextBoxSubject);
            ToolTipUtil.SetDelays(ref m_tool_tip_subject);
            m_tool_tip_text.SetToolTip(this.m_rich_textbox_body, NewsLetterSettings.Default.ToolTipTextBoxText);
            ToolTipUtil.SetDelays(ref m_tool_tip_text);
            m_tool_tip_picture_box.SetToolTip(this.m_picture_box, NewsLetterSettings.Default.ToolTipPictureBox);
            ToolTipUtil.SetDelays(ref m_tool_tip_picture_box);
            m_tool_tip_test_send.SetToolTip(this.m_button_test_send, NewsLetterSettings.Default.ToolTipButtonSendTest);
            ToolTipUtil.SetDelays(ref m_tool_tip_test_send);
            m_tool_tip_send_all.SetToolTip(this.m_button_all_send, NewsLetterSettings.Default.ToolTipButtonSendAll);
            ToolTipUtil.SetDelays(ref m_tool_tip_send_all);
            m_tool_tip_text_box_to.SetToolTip(this.m_textbox_to, NewsLetterSettings.Default.ToolTipTextBoxTo);
            ToolTipUtil.SetDelays(ref m_tool_tip_text_box_to);
            m_tool_tip_combo_box_from.SetToolTip(this.m_combo_box_from, NewsLetterSettings.Default.ToolTipComboBoxFrom);
            ToolTipUtil.SetDelays(ref m_tool_tip_combo_box_from);
            m_tool_tip_combo_box_test.SetToolTip(this.m_combo_box_bcc_test, NewsLetterSettings.Default.ToolTipComboBoxTestAddress);
            ToolTipUtil.SetDelays(ref m_tool_tip_combo_box_test);
            m_tool_tip_end_session.SetToolTip(this.m_button_ende, NewsLetterSettings.Default.ToolTipButtonEnd);
            ToolTipUtil.SetDelays(ref m_tool_tip_end_session);
            m_tool_tip_attachment.SetToolTip(this.m_combo_box_attachment, NewsLetterSettings.Default.ToolTipComboBoxAttachment);
            ToolTipUtil.SetDelays(ref m_tool_tip_attachment);
            m_tool_tip_edit_config.SetToolTip(this.m_button_edit_config, NewsLetterSettings.Default.ToolTipButtonEditConfig);
            ToolTipUtil.SetDelays(ref m_tool_tip_edit_config);
            m_tool_tip_help.SetToolTip(this.m_button_help, NewsLetterSettings.Default.ToolTipButtonHelp);
            ToolTipUtil.SetDelays(ref m_tool_tip_help);
            m_tool_tip_update.SetToolTip(this.m_button_update, NewsLetterSettings.Default.ToolTipButtonUpdate);
            ToolTipUtil.SetDelays(ref m_tool_tip_update);
            m_tool_tip_only_downloaded_pics.SetToolTip(this.m_checkbox_only_down_loaded_pics, NewsLetterSettings.Default.ToolTipCheckBoxPics);
            ToolTipUtil.SetDelays(ref m_tool_tip_only_downloaded_pics);
            m_tool_tip_checkbox_with_reservation_text.SetToolTip(this.m_checkbox_with_reservation_text, NewsLetterSettings.Default.ToolTipCheckBoxAddReservation);
            ToolTipUtil.SetDelays(ref m_tool_tip_checkbox_with_reservation_text);

        } //  _SetToolTips()

        #region Set controls
        /// <summary>Initialization of controls</summary>
        private void _InitializeControls()
        {
            string error_message = @"";

            if (!NewsLetterXml.SetDateTimePickerToNextConcertDate(m_date_time_picker, out error_message))
            {
                MessageBox.Show(error_message);

                return;
            }

            //this.m_button_edit_config.Enabled = false;
            //if (NewsLetterSettings.Default.DoNotSendAll)
            {
                this.m_button_edit_config.Enabled = true;
            }

            this.m_update_subject = true;

            _SetControlCaptions();

            _SetTimeComboBoxes();

            _SetFromComboBox();

            _SetTestAddressComboBox();

            _SetPicsCombo();

            _SetCheckBoxDownloadedPosters();

            _SetAttachmentCombo();

            _SetSubjectText();
        }

        /// <summary>Set the captions for the controls</summary>
        private void _SetControlCaptions()
        {
            this.Text = NewsLetterSettings.Default.GuiTextProgramTitle;
            if (NewsLetterSettings.Default.DoNotSendAll)
            {
                this.Text = NewsLetterSettings.Default.GuiTextProgramTitleTest;
            }
            this.m_button_ende.Text = NewsLetterSettings.Default.GuiButtonEnd;
            this.m_text_box_band.Text = NewsLetterSettings.Default.GuiTextBand;
            this.m_textbox_to.Text = NewsLetterSettings.Default.EmailAddressInfo;
            this.m_button_subject_combine.Text = NewsLetterSettings.Default.GuiButtonAutomatic;
            this.m_button_all_send.Text = NewsLetterSettings.Default.GuiButtonSendToAll;
            this.m_button_test_send.Text = NewsLetterSettings.Default.GuiButtonSendTest;
            this.m_button_edit_config.Text = NewsLetterSettings.Default.GuiButtonEditConfig;
            this.m_button_help.Text = NewsLetterSettings.Default.GuiButtonHelp;
            this.m_button_update.Text = NewsLetterSettings.Default.GuiButtonUpdate;

            this.m_label_test_address.Text = NewsLetterSettings.Default.GuiLabelTestAddress;
            this.m_label_from.Text = NewsLetterSettings.Default.GuiLabelFrom;
            this.m_label_to.Text = NewsLetterSettings.Default.GuiLabelTo;
            this.labelText.Text = NewsLetterSettings.Default.GuiLabelText;
            this.m_label_poster.Text = NewsLetterSettings.Default.GuiLabelPoster;
            this.m_label_band.Text = NewsLetterSettings.Default.GuiLabelBand;
            this.m_label_zeit.Text = NewsLetterSettings.Default.GuiLabelTime;
            this.m_label_datum.Text = NewsLetterSettings.Default.GuiLabelDate;
            this.m_label_attachment.Text = NewsLetterSettings.Default.GuiLabelAttachment;
            this.m_label_subject.Text = NewsLetterSettings.Default.GuiLabelSubject;
            this.m_label_version.Text = NewsLetterSettings.Default.GuiLabelVersion;
            this.m_rich_textbox_body.Text = "";
            this.m_textbox_message.Text = " ";
            this.m_checkbox_only_down_loaded_pics.Text = NewsLetterSettings.Default.GuiLabelShowOnlyDownloadedPics;
            this.m_checkbox_with_reservation_text.Text = NewsLetterSettings.Default.GuiLabelAddReservationText;
        }

        /// <summary>Set time comboboxes</summary>
        private void _SetTimeComboBoxes()
        {
            for (int i_hour = 0; i_hour <= 23; i_hour++)
            {
                string current_hour = i_hour.ToString();
                if (i_hour < 10)
                {
                    current_hour = "0" + current_hour;
                }
                current_hour = " " + current_hour;

                m_combo_box_hour.Items.Add(current_hour);
            }

            m_combo_box_hour.Text = " 15"; // Default

            for (int i_minute = 0; i_minute <= 59; i_minute++)
            {
                string current_minute = i_minute.ToString();
                if (i_minute < 10)
                {
                    current_minute = "0" + current_minute;
                }

                m_combo_box_minute.Items.Add(current_minute);
            }

            m_combo_box_minute.Text = "30"; // Default
        }

        /// <summary>Set the From combobox</summary>
        private void _SetFromComboBox()
        {
            string error_message = @"";

            if (!NewsLetterXml.SetFromComboBox(m_combo_box_from, out error_message))
            {
                error_message = @"NewsletterForm._SetFromComboBox NewsLetterXml.SetFromComboBox failed " + error_message;

                MessageBox.Show(error_message);
                return;
            }

        } // _SetFromComboBox

        /// <summary>Set the test address combobox</summary>
        private void _SetTestAddressComboBox()
        {
            string error_message = @"";

            if (!NewsLetterXml.SetTestAddressComboBox(m_combo_box_bcc_test, out error_message))
            {
                error_message = @"NewsletterForm._SetFromComboBox NewsLetterXml.SetTestAddressComboBox failed " + error_message;

                MessageBox.Show(error_message);
                return;
            }

        } // _SetTestAddressComboBox

        /// <summary>Set check box for flag that tells if only downloaded posters shall be shown</summary>
        private void _SetCheckBoxDownloadedPosters()
        {
            if (NewsLetterSettings.Default.DownloadedPostersOnly)
            {
                this.m_checkbox_only_down_loaded_pics.Checked = true;
            }
            else
            {
                this.m_checkbox_only_down_loaded_pics.Checked = false;
            }
        }

        /// <summary>Set combobox for pictures</summary>
        private void _SetPicsCombo()
        {
            string poster_directory = FileUtil.PosterDirectory();

            string[] file_extensions;

            ArrayList file_extensions_string_array = new ArrayList();
            file_extensions_string_array.Add(".bmp");
            file_extensions_string_array.Add(".png");
            file_extensions_string_array.Add(".gif");
            file_extensions_string_array.Add(".jpg");

            file_extensions = (string[])file_extensions_string_array.ToArray(typeof(string));

            bool b_get = FileUtil.GetFilesDirectory(file_extensions, poster_directory, out m_picture_file_names);

            m_combo_box_pics.Items.Clear();
            if (!b_get || m_picture_file_names.Length == 0)
            {
                m_combo_box_pics.Items.Add("");
                m_combo_box_pics.Text = "";

                return;
            }

            bool downloaded_posters_only = this.m_checkbox_only_down_loaded_pics.Checked;

            foreach (string file_name in m_picture_file_names)
            {
                string file_name_without_path = Path.GetFileName(file_name);
                bool b_downloaded = FileUtil.FileIsInArray(file_name_without_path, m_downloaded_picture_file_names);
                if (downloaded_posters_only)
                {
                    if (b_downloaded)
                    {
                        m_combo_box_pics.Items.Add(file_name_without_path);
                    }
                }
                else
                {
                    m_combo_box_pics.Items.Add(file_name_without_path);
                }
            }

            m_combo_box_pics.Items.Add(""); // Last empty line when no picture shall be sent

            string file_name_default = Path.GetFileName(m_picture_file_names[m_picture_file_names.Length - 1]);
            //m_combo_box_pics.Text = file_name_default;

            m_combo_box_pics.Text = "";

            NewsLetterXml.SetPosterToNextConcert(m_combo_box_pics);

        } // _SetPicsCombo

        /// <summary>Set combobox for attachments</summary>
        private void _SetAttachmentCombo()
        {
            string attachment_directory = FileUtil.AttachmentDirectory();

            string[] file_extensions;

            ArrayList file_extensions_string_array = new ArrayList();
            file_extensions_string_array.Add(".bmp");
            file_extensions_string_array.Add(".png");
            file_extensions_string_array.Add(".gif");
            file_extensions_string_array.Add(".jpg");
            file_extensions_string_array.Add(".doc");
            file_extensions_string_array.Add(".docx");
            file_extensions_string_array.Add(".txt");
            file_extensions_string_array.Add(".pdf");

            file_extensions = (string[])file_extensions_string_array.ToArray(typeof(string));

            bool b_get = FileUtil.GetFilesDirectory(file_extensions, attachment_directory, out m_attachment_file_names);

            if (!b_get || m_attachment_file_names.Length == 0)
            {
                m_combo_box_attachment.Items.Add("");
                m_combo_box_attachment.Text = "";

                return;
            }

            m_combo_box_attachment.Items.Clear();
            foreach (string file_name in m_attachment_file_names)
            {
                string file_name_without_path = Path.GetFileName(file_name);
                m_combo_box_attachment.Items.Add(file_name_without_path);
            }

            m_combo_box_attachment.Items.Add(""); // Last empty line when no attachment shall be sent

            string file_name_default = Path.GetFileName(m_attachment_file_names[m_attachment_file_names.Length - 1]);
            //m_combo_box_attachment.Text = file_name_default;

            m_combo_box_attachment.Text = "";
        }


        /// <summary>Combine the subject text from other controls</summary>
        private void _SetSubjectText()
        {
            if (!this.m_update_subject)
            {
                // User has made a manual change in the subject text
                return;
            }

            DateTime set_value = this.m_date_time_picker.Value;
            int set_year = set_value.Year;
            int set_month = set_value.Month;
            int set_day = set_value.Day;
            string str_date = set_day.ToString() + @"/" + set_month.ToString() + @" " + set_year.ToString();

            string str_hour = m_combo_box_hour.Text;
            if (str_hour.StartsWith("0"))
            {
                str_hour = str_hour.Substring(1, 1);
            }

            string str_minute = m_combo_box_minute.Text;

            this.m_textbox_subject.Text = NewsLetterSettings.Default.SubjectStartText +
                " " + str_date + " " + str_hour + ":" + str_minute + "  " +
                this.m_text_box_band.Text;

            this.m_update_subject = true;
            this.m_button_subject_combine.Text = NewsLetterSettings.Default.GuiButtonAutomatic;
        }

        #endregion

        #region TODO
        /*TODO
        private void _LogfileCreateNew(string i_header, out string o_log_file_name)
        {
            _GetOutputLogFileName(out o_log_file_name);

            FileStream fsOutput = new FileStream(o_log_file_name, FileMode.Create,
                FileAccess.Write);
            StreamWriter srOutput = new StreamWriter(fsOutput);

            srOutput.WriteLine(i_header);
            srOutput.WriteLine("File: " + o_log_file_name);
            srOutput.WriteLine("");

            srOutput.Close();
            fsOutput.Close();
        }

        private void _LogfileAppend(string i_text, string i_log_file_name)
        {
            FileStream fsOutput = new FileStream(i_log_file_name, FileMode.Append,
                FileAccess.Write);
            StreamWriter srOutput = new StreamWriter(fsOutput);

            srOutput.WriteLine(i_text);

            srOutput.Close();
            fsOutput.Close();
        }
         * TODO*/

        #endregion

        #region Send buttons

        /// <summary>Get input data from the dialog window and set Newsletter data object</summary>
        private bool _GetInputFromDialog(ref NewsletterData io_newsletter_data)
		{
			ResourceManager rm = new ResourceManager(typeof(NewsletterForm));

            if (this.m_combo_box_from.Text == "")
			{
                this.m_combo_box_from.Focus();
                MessageBox.Show(NewsLetterSettings.Default.ErrMsgNoFrom);
				return false;
			}
            

			if(this.m_textbox_to.Text == "")
			{
				this.m_textbox_to.Focus();
                MessageBox.Show(NewsLetterSettings.Default.ErrMsgNoTo);
                return false;
			}
            
			if(this.m_textbox_subject.Text == "")
			{
				m_textbox_subject.Focus();
                MessageBox.Show(NewsLetterSettings.Default.ErrMsgNoSubject);
                return false;
			}

            io_newsletter_data.m_from = this.m_combo_box_from.Text;
            io_newsletter_data.m_to = this.m_textbox_to.Text;
            io_newsletter_data.m_subject = this.m_textbox_subject.Text;
            io_newsletter_data.m_bcc_test = this.m_combo_box_bcc_test.Text;

            io_newsletter_data.m_date = this.m_date_time_picker.Text;
            if (io_newsletter_data.m_date.Trim() == "")
                io_newsletter_data.m_date = "?/?";
            int index_day_0 = io_newsletter_data.m_date.IndexOf("0");
            if (0 == index_day_0)
            {
                io_newsletter_data.m_date = io_newsletter_data.m_date.Substring(1);
            }
            int index_month_0 = io_newsletter_data.m_date.IndexOf("/0");
            if (index_day_0 > 0)
            {
                io_newsletter_data.m_date = io_newsletter_data.m_date.Replace("/0", "/");
            }

            io_newsletter_data.m_band = this.m_text_box_band.Text;

            io_newsletter_data.AddReservationText = this.m_checkbox_with_reservation_text.Checked;

            string error_message = "";
            if (!io_newsletter_data.SetPosterName(this.m_combo_box_pics.Text, out error_message))
            {
                m_combo_box_pics.Focus();
                MessageBox.Show(error_message);
                return false;
            }

            if (!io_newsletter_data.SetAttachmentName(this.m_combo_box_attachment.Text, out error_message))
            {
                m_combo_box_attachment.Focus();
                MessageBox.Show(error_message);
                return false;
            }

            if (!io_newsletter_data.SetBodyText(this.m_rich_textbox_body.Text, out error_message))
            {
                m_textbox_subject.Focus();
                MessageBox.Show(error_message);
                return false;
            }

            if (!io_newsletter_data.Check(out error_message))
            {
                m_textbox_subject.Focus();
                MessageBox.Show(error_message);
                return false;
            }

            return true;
		}

        /// <summary>Send test newsletter(s)</summary>
        private void _ButtonSendTest_Click(object sender, EventArgs e)
        {
            NewsletterData newsletter_data = new NewsletterData();

            if (!_GetInputFromDialog(ref newsletter_data))
                return;

            string error_message;
           
            if (!this.m_newsletter_main.SendNewsLetterTest(newsletter_data, out error_message))
            {
                MessageBox.Show(error_message, " ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return;
            }


        } // _ButtonSendTest_Click

        /// <summary>Send newsletters to all in the .csv file that has subscribed to the newsletter</summary>
        private void _ButtonSendAll_Click(object sender, System.EventArgs e)
        {
            NewsletterData newsletter_data = new NewsletterData();

            if (!_GetInputFromDialog(ref newsletter_data))
                return;

            string error_message;
            if (!this.m_newsletter_main.SendNewsLetterAll(newsletter_data, out error_message))
            {
                MessageBox.Show(error_message, " ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return;
            }

        } // _ButtonSendAll_Click

        /// <summary>Download new versions, posters and attachments from the Web server</summary>
        private void m_button_update_Click(object sender, EventArgs e)
        {
            bool b_new_version = true;
            string error_message = "";
            if (!m_newsletter_main.DownloadFiles(b_new_version, out error_message))
            {
                MessageBox.Show(error_message);
            }
            else
            {
                MessageBox.Show(NewsLetterSettings.Default.MsgFilesDownload);
            }

            _SetPicsCombo();
            _SetAttachmentCombo();
        }

 
        #endregion

        #region Event handlers
        /// <summary>Event: New date</summary>
        private void m_date_time_picker_ValueChanged(object sender, EventArgs e)
        {
            _SetSubjectText();
        }

        /// <summary>Event: New hour</summary>
        private void m_combo_box_hour_SelectedIndexChanged(object sender, EventArgs e)
        {
            _SetSubjectText();
        }

        /// <summary>Event: New minute</summary>
        private void m_combo_box_minute_SelectedIndexChanged(object sender, EventArgs e)
        {
            _SetSubjectText();
        }

        /// <summary>Event: Band text changed</summary>
        private void m_text_box_band_TextChanged(object sender, EventArgs e)
        {
            _SetSubjectText();
        }

        /// <summary>Event: Combine text automatically from other controls</summary>
        private void m_button_subject_combine_Click(object sender, EventArgs e)
        {
            this.m_update_subject = true;
            this.m_button_subject_combine.Text = NewsLetterSettings.Default.GuiButtonAutomatic;

            _SetSubjectText();
        }

        /// <summary>Event: Flag telling if only downloaded pictures shall be shown has been changed</summary>
        private void m_checkbox_only_down_loaded_pics_CheckedChanged(object sender, EventArgs e)
        {
            _SetPicsCombo();
        }

        /// <summary>Event: The configuration file has been changed. Read the new values from the config file and update the controls</summary>
        public void ConfigFileHasBeenChanged()
        {
            NewsLetterSettings.Default.ReadFromConfigFile();

            _SetControlCaptions();

            _SetToolTips();

            _SetFromComboBox();

            _SetTestAddressComboBox();
        }

        /// <summary>Event: Help</summary>
        private void m_button_help_Click(object sender, EventArgs e)
        {
            HelpForm help_form = new HelpForm();
            help_form.Owner = this;
            help_form.ShowDialog();
        }

        /// <summary>Event: Edit configuration file</summary>
        private void m_button_edit_config_Click(object sender, EventArgs e)
        {
            EditConfigFileForm form_settings = new EditConfigFileForm(this);
            form_settings.Owner = this;
            form_settings.ShowDialog();
        }

        /// <summary>Event: Write text manually</summary>
        private void m_textbox_subject_TextChanged(object sender, EventArgs e)
        {
            this.m_update_subject = false;
            this.m_button_subject_combine.Text = NewsLetterSettings.Default.GuiButtonManual;
        }

        /// <summary>Event: Poster is selected</summary>
        private void m_combo_box_pics_SelectedIndexChanged(object sender, EventArgs e)
        {
            string picture_file_name_with_path = "";
            string selected_picture = m_combo_box_pics.Text;
            selected_picture = selected_picture.Trim();
            if ("" == selected_picture)
            {
                _DisplayPoster(picture_file_name_with_path, false);
                return;
            }

            int pic_index = m_combo_box_pics.SelectedIndex;

            picture_file_name_with_path = m_picture_file_names[pic_index];

            if (!File.Exists(picture_file_name_with_path))
            {
                _DisplayPoster(picture_file_name_with_path, false);
                return;
            }

             _DisplayPoster(picture_file_name_with_path, true);

        }

        /// <summary>Displays poster/picture</summary>
        private void _DisplayPoster(string i_picture_file_name_with_path, bool i_picture_exist)
        {
            int pic_width = m_picture_box.Size.Width;
            int pic_height = m_picture_box.Size.Height;

            // Sets up an image object to be displayed.
            if (m_bitmap_poster != null)
            {
                m_bitmap_poster.Dispose();
            }

            // Stretches the image to fit the pictureBox.
            m_picture_box.SizeMode = PictureBoxSizeMode.StretchImage;
            if (i_picture_exist)
            {
                m_bitmap_poster = new Bitmap(i_picture_file_name_with_path);
            }
            else
            {
                m_bitmap_poster = Properties.Resources.PosterWithoutPicture;
            }

            m_picture_box.ClientSize = new Size(pic_width, pic_height);
            m_picture_box.Image = (Image)m_bitmap_poster;
        }

        /// <summary>Event: Attachment is selected</summary>
        private void m_combo_box_attachment_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        /// <summary>Remove highlightning of text in the combo boxes</summary>
        private void NewsletterForm_Paint(object sender, PaintEventArgs e)
        {
            // https://www.codeproject.com/Questions/754776/How-to-disable-highlighting-of-a-selected-item-in
            this.m_combo_box_hour.SelectionLength = 0;
            this.m_combo_box_minute.SelectionLength = 0;
            this.m_combo_box_pics.SelectionLength = 0;
            this.m_combo_box_attachment.SelectionLength = 0;
            this.m_combo_box_from.SelectionLength = 0;
            this.m_combo_box_bcc_test.SelectionLength = 0;

        } // NewsletterForm_Paint

        /// <summary>Event: Exit application</summary>
        private void _ButtonEnd_Click(object sender, System.EventArgs e)
        {
            Application.Exit();
        }

        #endregion

        #region Not used functions
        private void miPaste_Click(object sender, System.EventArgs e)
        {
            if (Clipboard.ContainsText(TextDataFormat.Rtf)) 
            { 
                this.m_rich_textbox_body.SelectedRtf = Clipboard.GetData(DataFormats.Rtf).ToString(); 
            }

        }

        private void rtbBody_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.cmExtras.Show(m_rich_textbox_body, new Point(e.X, e.Y));
            }
        }

        private void MainWnd_Load(object sender, System.EventArgs e)
        {

        }
       #endregion

        #region Basic functions
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>The main entry point for the application.</summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new NewsletterForm());
        }
        #endregion

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NewsletterForm));
            this.m_textbox_to = new System.Windows.Forms.TextBox();
            this.m_button_all_send = new System.Windows.Forms.Button();
            this.m_label_to = new System.Windows.Forms.Label();
            this.m_label_subject = new System.Windows.Forms.Label();
            this.m_textbox_subject = new System.Windows.Forms.TextBox();
            this.m_rich_textbox_body = new System.Windows.Forms.RichTextBox();
            this.mMainMenu = new System.Windows.Forms.MainMenu(this.components);
            this.cmExtras = new System.Windows.Forms.ContextMenu();
            this.cmiPaste = new System.Windows.Forms.MenuItem();
            this.cmiClearAll = new System.Windows.Forms.MenuItem();
            this.ImageList = new System.Windows.Forms.ImageList(this.components);
            this.m_label_from = new System.Windows.Forms.Label();
            this.m_label_poster = new System.Windows.Forms.Label();
            this.m_picture_logo = new System.Windows.Forms.PictureBox();
            this.m_button_ende = new System.Windows.Forms.Button();
            this.m_textbox_message = new System.Windows.Forms.TextBox();
            this.m_label_test_address = new System.Windows.Forms.Label();
            this.m_button_test_send = new System.Windows.Forms.Button();
            this.labelText = new System.Windows.Forms.Label();
            this.m_date_time_picker = new System.Windows.Forms.DateTimePicker();
            this.m_combo_box_hour = new System.Windows.Forms.ComboBox();
            this.m_combo_box_minute = new System.Windows.Forms.ComboBox();
            this.m_label_datum = new System.Windows.Forms.Label();
            this.m_label_zeit = new System.Windows.Forms.Label();
            this.m_label_colone = new System.Windows.Forms.Label();
            this.m_text_box_band = new System.Windows.Forms.TextBox();
            this.m_label_band = new System.Windows.Forms.Label();
            this.m_combo_box_from = new System.Windows.Forms.ComboBox();
            this.m_button_subject_combine = new System.Windows.Forms.Button();
            this.m_combo_box_pics = new System.Windows.Forms.ComboBox();
            this.m_combo_box_bcc_test = new System.Windows.Forms.ComboBox();
            this.m_picture_box = new System.Windows.Forms.PictureBox();
            this.m_rich_text_box_black = new System.Windows.Forms.RichTextBox();
            this.m_combo_box_attachment = new System.Windows.Forms.ComboBox();
            this.m_label_attachment = new System.Windows.Forms.Label();
            this.m_button_help = new System.Windows.Forms.Button();
            this.m_button_edit_config = new System.Windows.Forms.Button();
            this.m_button_update = new System.Windows.Forms.Button();
            this.m_checkbox_only_down_loaded_pics = new System.Windows.Forms.CheckBox();
            this.m_label_version = new System.Windows.Forms.Label();
            this.m_checkbox_with_reservation_text = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.m_picture_logo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_picture_box)).BeginInit();
            this.SuspendLayout();
            // 
            // m_textbox_to
            // 
            this.m_textbox_to.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.m_textbox_to.BackColor = System.Drawing.Color.Black;
            this.m_textbox_to.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_textbox_to.ForeColor = System.Drawing.Color.Red;
            this.m_textbox_to.Location = new System.Drawing.Point(264, 553);
            this.m_textbox_to.Name = "m_textbox_to";
            this.m_textbox_to.Size = new System.Drawing.Size(165, 22);
            this.m_textbox_to.TabIndex = 26;
            // 
            // m_button_all_send
            // 
            this.m_button_all_send.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_all_send.BackColor = System.Drawing.Color.Black;
            this.m_button_all_send.FlatAppearance.BorderColor = System.Drawing.Color.Red;
            this.m_button_all_send.FlatAppearance.BorderSize = 2;
            this.m_button_all_send.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.m_button_all_send.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_all_send.ForeColor = System.Drawing.Color.Red;
            this.m_button_all_send.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.m_button_all_send.Location = new System.Drawing.Point(595, 641);
            this.m_button_all_send.Name = "m_button_all_send";
            this.m_button_all_send.Size = new System.Drawing.Size(234, 29);
            this.m_button_all_send.TabIndex = 32;
            this.m_button_all_send.Text = "Send all newsletters";
            this.m_button_all_send.UseVisualStyleBackColor = false;
            this.m_button_all_send.Click += new System.EventHandler(this._ButtonSendAll_Click);
            // 
            // m_label_to
            // 
            this.m_label_to.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_label_to.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_to.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.m_label_to.Location = new System.Drawing.Point(275, 535);
            this.m_label_to.Name = "m_label_to";
            this.m_label_to.Size = new System.Drawing.Size(434, 15);
            this.m_label_to.TabIndex = 24;
            this.m_label_to.Text = "To";
            // 
            // m_label_subject
            // 
            this.m_label_subject.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_subject.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.m_label_subject.Location = new System.Drawing.Point(17, 112);
            this.m_label_subject.Name = "m_label_subject";
            this.m_label_subject.Size = new System.Drawing.Size(118, 17);
            this.m_label_subject.TabIndex = 13;
            this.m_label_subject.Text = "Subject";
            // 
            // m_textbox_subject
            // 
            this.m_textbox_subject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_textbox_subject.BackColor = System.Drawing.Color.Black;
            this.m_textbox_subject.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_textbox_subject.ForeColor = System.Drawing.Color.Red;
            this.m_textbox_subject.Location = new System.Drawing.Point(9, 133);
            this.m_textbox_subject.Name = "m_textbox_subject";
            this.m_textbox_subject.Size = new System.Drawing.Size(889, 22);
            this.m_textbox_subject.TabIndex = 14;
            this.m_textbox_subject.TextChanged += new System.EventHandler(this.m_textbox_subject_TextChanged);
            // 
            // m_rich_textbox_body
            // 
            this.m_rich_textbox_body.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_rich_textbox_body.BackColor = System.Drawing.Color.White;
            this.m_rich_textbox_body.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_rich_textbox_body.ForeColor = System.Drawing.Color.Black;
            this.m_rich_textbox_body.Location = new System.Drawing.Point(264, 191);
            this.m_rich_textbox_body.Name = "m_rich_textbox_body";
            this.m_rich_textbox_body.Size = new System.Drawing.Size(662, 288);
            this.m_rich_textbox_body.TabIndex = 21;
            this.m_rich_textbox_body.Text = "";
            this.m_rich_textbox_body.MouseDown += new System.Windows.Forms.MouseEventHandler(this.rtbBody_MouseDown);
            // 
            // cmExtras
            // 
            this.cmExtras.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.cmiPaste});
            // 
            // cmiPaste
            // 
            this.cmiPaste.Index = 0;
            this.cmiPaste.Text = "Paste";
            this.cmiPaste.Click += new System.EventHandler(this.miPaste_Click);
            // 
            // cmiClearAll
            // 
            this.cmiClearAll.Index = -1;
            this.cmiClearAll.Text = "";
            // 
            // ImageList
            // 
            this.ImageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ImageList.ImageStream")));
            this.ImageList.TransparentColor = System.Drawing.Color.Transparent;
            this.ImageList.Images.SetKeyName(0, "");
            this.ImageList.Images.SetKeyName(1, "");
            // 
            // m_label_from
            // 
            this.m_label_from.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_label_from.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_from.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.m_label_from.Location = new System.Drawing.Point(456, 535);
            this.m_label_from.Name = "m_label_from";
            this.m_label_from.Size = new System.Drawing.Size(434, 19);
            this.m_label_from.TabIndex = 25;
            this.m_label_from.Text = "From";
            // 
            // m_label_poster
            // 
            this.m_label_poster.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_poster.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.m_label_poster.Location = new System.Drawing.Point(16, 167);
            this.m_label_poster.Name = "m_label_poster";
            this.m_label_poster.Size = new System.Drawing.Size(48, 16);
            this.m_label_poster.TabIndex = 16;
            this.m_label_poster.Text = "Poster";
            // 
            // m_picture_logo
            // 
            this.m_picture_logo.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.m_picture_logo.BackColor = System.Drawing.Color.Black;
            this.m_picture_logo.Image = ((System.Drawing.Image)(resources.GetObject("m_picture_logo.Image")));
            this.m_picture_logo.Location = new System.Drawing.Point(243, 1);
            this.m_picture_logo.Name = "m_picture_logo";
            this.m_picture_logo.Size = new System.Drawing.Size(287, 39);
            this.m_picture_logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.m_picture_logo.TabIndex = 15;
            this.m_picture_logo.TabStop = false;
            // 
            // m_button_ende
            // 
            this.m_button_ende.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_ende.BackColor = System.Drawing.Color.Black;
            this.m_button_ende.FlatAppearance.BorderColor = System.Drawing.Color.Red;
            this.m_button_ende.FlatAppearance.BorderSize = 2;
            this.m_button_ende.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.m_button_ende.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_ende.Location = new System.Drawing.Point(865, 681);
            this.m_button_ende.Name = "m_button_ende";
            this.m_button_ende.Size = new System.Drawing.Size(63, 27);
            this.m_button_ende.TabIndex = 33;
            this.m_button_ende.Text = "End";
            this.m_button_ende.UseVisualStyleBackColor = false;
            this.m_button_ende.Click += new System.EventHandler(this._ButtonEnd_Click);
            // 
            // m_textbox_message
            // 
            this.m_textbox_message.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_textbox_message.BackColor = System.Drawing.Color.Black;
            this.m_textbox_message.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.m_textbox_message.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_textbox_message.ForeColor = System.Drawing.Color.Red;
            this.m_textbox_message.Location = new System.Drawing.Point(12, 688);
            this.m_textbox_message.Name = "m_textbox_message";
            this.m_textbox_message.Size = new System.Drawing.Size(841, 15);
            this.m_textbox_message.TabIndex = 31;
            // 
            // m_label_test_address
            // 
            this.m_label_test_address.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.m_label_test_address.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_test_address.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.m_label_test_address.Location = new System.Drawing.Point(275, 582);
            this.m_label_test_address.Name = "m_label_test_address";
            this.m_label_test_address.Size = new System.Drawing.Size(210, 16);
            this.m_label_test_address.TabIndex = 28;
            this.m_label_test_address.Text = "Test address";
            // 
            // m_button_test_send
            // 
            this.m_button_test_send.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_test_send.BackColor = System.Drawing.Color.Black;
            this.m_button_test_send.FlatAppearance.BorderColor = System.Drawing.Color.Red;
            this.m_button_test_send.FlatAppearance.BorderSize = 2;
            this.m_button_test_send.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.m_button_test_send.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_test_send.Location = new System.Drawing.Point(733, 600);
            this.m_button_test_send.Name = "m_button_test_send";
            this.m_button_test_send.Size = new System.Drawing.Size(198, 27);
            this.m_button_test_send.TabIndex = 30;
            this.m_button_test_send.Text = "Send a test newsletter";
            this.m_button_test_send.UseVisualStyleBackColor = false;
            this.m_button_test_send.Click += new System.EventHandler(this._ButtonSendTest_Click);
            // 
            // labelText
            // 
            this.labelText.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelText.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.labelText.Location = new System.Drawing.Point(275, 166);
            this.labelText.Name = "labelText";
            this.labelText.Size = new System.Drawing.Size(48, 16);
            this.labelText.TabIndex = 18;
            this.labelText.Text = "Text";
            // 
            // m_date_time_picker
            // 
            this.m_date_time_picker.CalendarFont = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_date_time_picker.CalendarForeColor = System.Drawing.Color.Red;
            this.m_date_time_picker.CalendarTitleForeColor = System.Drawing.Color.Red;
            this.m_date_time_picker.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_date_time_picker.Location = new System.Drawing.Point(9, 79);
            this.m_date_time_picker.Name = "m_date_time_picker";
            this.m_date_time_picker.Size = new System.Drawing.Size(56, 22);
            this.m_date_time_picker.TabIndex = 8;
            this.m_date_time_picker.ValueChanged += new System.EventHandler(this.m_date_time_picker_ValueChanged);
            // 
            // m_combo_box_hour
            // 
            this.m_combo_box_hour.BackColor = System.Drawing.Color.Black;
            this.m_combo_box_hour.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_hour.ForeColor = System.Drawing.Color.Red;
            this.m_combo_box_hour.FormattingEnabled = true;
            this.m_combo_box_hour.Location = new System.Drawing.Point(78, 79);
            this.m_combo_box_hour.Name = "m_combo_box_hour";
            this.m_combo_box_hour.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.m_combo_box_hour.Size = new System.Drawing.Size(42, 24);
            this.m_combo_box_hour.TabIndex = 9;
            this.m_combo_box_hour.SelectedIndexChanged += new System.EventHandler(this.m_combo_box_hour_SelectedIndexChanged);
            // 
            // m_combo_box_minute
            // 
            this.m_combo_box_minute.BackColor = System.Drawing.Color.Black;
            this.m_combo_box_minute.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_minute.ForeColor = System.Drawing.Color.Red;
            this.m_combo_box_minute.FormattingEnabled = true;
            this.m_combo_box_minute.Location = new System.Drawing.Point(131, 79);
            this.m_combo_box_minute.Name = "m_combo_box_minute";
            this.m_combo_box_minute.Size = new System.Drawing.Size(42, 24);
            this.m_combo_box_minute.TabIndex = 10;
            this.m_combo_box_minute.SelectedIndexChanged += new System.EventHandler(this.m_combo_box_minute_SelectedIndexChanged);
            // 
            // m_label_datum
            // 
            this.m_label_datum.AutoSize = true;
            this.m_label_datum.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_datum.Location = new System.Drawing.Point(17, 59);
            this.m_label_datum.Name = "m_label_datum";
            this.m_label_datum.Size = new System.Drawing.Size(37, 16);
            this.m_label_datum.TabIndex = 5;
            this.m_label_datum.Text = "Date";
            // 
            // m_label_zeit
            // 
            this.m_label_zeit.AutoSize = true;
            this.m_label_zeit.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_zeit.Location = new System.Drawing.Point(105, 59);
            this.m_label_zeit.Name = "m_label_zeit";
            this.m_label_zeit.Size = new System.Drawing.Size(40, 16);
            this.m_label_zeit.TabIndex = 6;
            this.m_label_zeit.Text = "Time";
            // 
            // m_label_colone
            // 
            this.m_label_colone.AutoSize = true;
            this.m_label_colone.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_colone.Location = new System.Drawing.Point(121, 80);
            this.m_label_colone.Name = "m_label_colone";
            this.m_label_colone.Size = new System.Drawing.Size(12, 16);
            this.m_label_colone.TabIndex = 11;
            this.m_label_colone.Text = ":";
            // 
            // m_text_box_band
            // 
            this.m_text_box_band.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_text_box_band.BackColor = System.Drawing.Color.Black;
            this.m_text_box_band.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_text_box_band.ForeColor = System.Drawing.Color.Red;
            this.m_text_box_band.Location = new System.Drawing.Point(185, 81);
            this.m_text_box_band.Name = "m_text_box_band";
            this.m_text_box_band.Size = new System.Drawing.Size(746, 22);
            this.m_text_box_band.TabIndex = 12;
            this.m_text_box_band.TextChanged += new System.EventHandler(this.m_text_box_band_TextChanged);
            // 
            // m_label_band
            // 
            this.m_label_band.AutoSize = true;
            this.m_label_band.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_band.Location = new System.Drawing.Point(196, 59);
            this.m_label_band.Name = "m_label_band";
            this.m_label_band.Size = new System.Drawing.Size(41, 16);
            this.m_label_band.TabIndex = 7;
            this.m_label_band.Text = "Band";
            // 
            // m_combo_box_from
            // 
            this.m_combo_box_from.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_combo_box_from.BackColor = System.Drawing.Color.Black;
            this.m_combo_box_from.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_from.ForeColor = System.Drawing.Color.Red;
            this.m_combo_box_from.FormattingEnabled = true;
            this.m_combo_box_from.Location = new System.Drawing.Point(443, 553);
            this.m_combo_box_from.Margin = new System.Windows.Forms.Padding(0);
            this.m_combo_box_from.Name = "m_combo_box_from";
            this.m_combo_box_from.Size = new System.Drawing.Size(486, 24);
            this.m_combo_box_from.TabIndex = 27;
            // 
            // m_button_subject_combine
            // 
            this.m_button_subject_combine.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_subject_combine.BackColor = System.Drawing.Color.Black;
            this.m_button_subject_combine.FlatAppearance.BorderColor = System.Drawing.Color.Red;
            this.m_button_subject_combine.FlatAppearance.BorderSize = 2;
            this.m_button_subject_combine.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.m_button_subject_combine.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_subject_combine.Location = new System.Drawing.Point(906, 132);
            this.m_button_subject_combine.Name = "m_button_subject_combine";
            this.m_button_subject_combine.Size = new System.Drawing.Size(24, 23);
            this.m_button_subject_combine.TabIndex = 15;
            this.m_button_subject_combine.Text = "U";
            this.m_button_subject_combine.UseVisualStyleBackColor = false;
            this.m_button_subject_combine.Click += new System.EventHandler(this.m_button_subject_combine_Click);
            // 
            // m_combo_box_pics
            // 
            this.m_combo_box_pics.BackColor = System.Drawing.Color.Black;
            this.m_combo_box_pics.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_pics.ForeColor = System.Drawing.Color.Red;
            this.m_combo_box_pics.FormattingEnabled = true;
            this.m_combo_box_pics.Location = new System.Drawing.Point(12, 191);
            this.m_combo_box_pics.Margin = new System.Windows.Forms.Padding(0);
            this.m_combo_box_pics.Name = "m_combo_box_pics";
            this.m_combo_box_pics.Size = new System.Drawing.Size(228, 24);
            this.m_combo_box_pics.TabIndex = 20;
            this.m_combo_box_pics.SelectedIndexChanged += new System.EventHandler(this.m_combo_box_pics_SelectedIndexChanged);
            // 
            // m_combo_box_bcc_test
            // 
            this.m_combo_box_bcc_test.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_combo_box_bcc_test.BackColor = System.Drawing.Color.Black;
            this.m_combo_box_bcc_test.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_bcc_test.ForeColor = System.Drawing.Color.Red;
            this.m_combo_box_bcc_test.FormattingEnabled = true;
            this.m_combo_box_bcc_test.Location = new System.Drawing.Point(264, 601);
            this.m_combo_box_bcc_test.Margin = new System.Windows.Forms.Padding(0);
            this.m_combo_box_bcc_test.Name = "m_combo_box_bcc_test";
            this.m_combo_box_bcc_test.Size = new System.Drawing.Size(466, 24);
            this.m_combo_box_bcc_test.TabIndex = 29;
            // 
            // m_picture_box
            // 
            this.m_picture_box.BackgroundImage = global::SendNewsLetter.Properties.Resources.PosterWithoutPicture;
            this.m_picture_box.Location = new System.Drawing.Point(19, 233);
            this.m_picture_box.Name = "m_picture_box";
            this.m_picture_box.Size = new System.Drawing.Size(218, 293);
            this.m_picture_box.TabIndex = 34;
            this.m_picture_box.TabStop = false;
            // 
            // m_rich_text_box_black
            // 
            this.m_rich_text_box_black.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_rich_text_box_black.BackColor = System.Drawing.Color.Black;
            this.m_rich_text_box_black.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.m_rich_text_box_black.Location = new System.Drawing.Point(5, 6);
            this.m_rich_text_box_black.Name = "m_rich_text_box_black";
            this.m_rich_text_box_black.Size = new System.Drawing.Size(933, 58);
            this.m_rich_text_box_black.TabIndex = 0;
            this.m_rich_text_box_black.Text = "";
            // 
            // m_combo_box_attachment
            // 
            this.m_combo_box_attachment.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_combo_box_attachment.BackColor = System.Drawing.Color.Black;
            this.m_combo_box_attachment.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_combo_box_attachment.ForeColor = System.Drawing.Color.Red;
            this.m_combo_box_attachment.FormattingEnabled = true;
            this.m_combo_box_attachment.Location = new System.Drawing.Point(264, 506);
            this.m_combo_box_attachment.Margin = new System.Windows.Forms.Padding(0);
            this.m_combo_box_attachment.Name = "m_combo_box_attachment";
            this.m_combo_box_attachment.Size = new System.Drawing.Size(666, 24);
            this.m_combo_box_attachment.TabIndex = 23;
            this.m_combo_box_attachment.SelectedIndexChanged += new System.EventHandler(this.m_combo_box_attachment_SelectedIndexChanged);
            // 
            // m_label_attachment
            // 
            this.m_label_attachment.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.m_label_attachment.AutoSize = true;
            this.m_label_attachment.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_attachment.Location = new System.Drawing.Point(275, 486);
            this.m_label_attachment.Name = "m_label_attachment";
            this.m_label_attachment.Size = new System.Drawing.Size(80, 16);
            this.m_label_attachment.TabIndex = 22;
            this.m_label_attachment.Text = "Attachment";
            // 
            // m_button_help
            // 
            this.m_button_help.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_help.BackColor = System.Drawing.SystemColors.WindowText;
            this.m_button_help.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_help.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.m_button_help.Location = new System.Drawing.Point(893, 5);
            this.m_button_help.Name = "m_button_help";
            this.m_button_help.Size = new System.Drawing.Size(40, 21);
            this.m_button_help.TabIndex = 3;
            this.m_button_help.Text = "Help";
            this.m_button_help.UseVisualStyleBackColor = false;
            this.m_button_help.Click += new System.EventHandler(this.m_button_help_Click);
            // 
            // m_button_edit_config
            // 
            this.m_button_edit_config.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_edit_config.BackColor = System.Drawing.SystemColors.WindowText;
            this.m_button_edit_config.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_edit_config.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.m_button_edit_config.Location = new System.Drawing.Point(778, 5);
            this.m_button_edit_config.Name = "m_button_edit_config";
            this.m_button_edit_config.Size = new System.Drawing.Size(40, 21);
            this.m_button_edit_config.TabIndex = 1;
            this.m_button_edit_config.Text = "Edit";
            this.m_button_edit_config.UseVisualStyleBackColor = false;
            this.m_button_edit_config.Click += new System.EventHandler(this.m_button_edit_config_Click);
            // 
            // m_button_update
            // 
            this.m_button_update.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_button_update.BackColor = System.Drawing.SystemColors.WindowText;
            this.m_button_update.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_button_update.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.m_button_update.Location = new System.Drawing.Point(819, 5);
            this.m_button_update.Name = "m_button_update";
            this.m_button_update.Size = new System.Drawing.Size(70, 21);
            this.m_button_update.TabIndex = 2;
            this.m_button_update.Text = "Download";
            this.m_button_update.UseVisualStyleBackColor = false;
            this.m_button_update.Click += new System.EventHandler(this.m_button_update_Click);
            // 
            // m_checkbox_only_down_loaded_pics
            // 
            this.m_checkbox_only_down_loaded_pics.AutoSize = true;
            this.m_checkbox_only_down_loaded_pics.Checked = true;
            this.m_checkbox_only_down_loaded_pics.CheckState = System.Windows.Forms.CheckState.Checked;
            this.m_checkbox_only_down_loaded_pics.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_checkbox_only_down_loaded_pics.Location = new System.Drawing.Point(91, 167);
            this.m_checkbox_only_down_loaded_pics.Name = "m_checkbox_only_down_loaded_pics";
            this.m_checkbox_only_down_loaded_pics.Size = new System.Drawing.Size(139, 20);
            this.m_checkbox_only_down_loaded_pics.TabIndex = 17;
            this.m_checkbox_only_down_loaded_pics.Text = "Only downloaded";
            this.m_checkbox_only_down_loaded_pics.UseVisualStyleBackColor = true;
            this.m_checkbox_only_down_loaded_pics.CheckedChanged += new System.EventHandler(this.m_checkbox_only_down_loaded_pics_CheckedChanged);
            // 
            // m_label_version
            // 
            this.m_label_version.AutoSize = true;
            this.m_label_version.BackColor = System.Drawing.Color.Black;
            this.m_label_version.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_label_version.ForeColor = System.Drawing.Color.Red;
            this.m_label_version.Location = new System.Drawing.Point(16, 34);
            this.m_label_version.Name = "m_label_version";
            this.m_label_version.Size = new System.Drawing.Size(69, 15);
            this.m_label_version.TabIndex = 4;
            this.m_label_version.Text = "Version 3.7";
            // 
            // m_checkbox_with_reservation_text
            // 
            this.m_checkbox_with_reservation_text.AutoSize = true;
            this.m_checkbox_with_reservation_text.Checked = true;
            this.m_checkbox_with_reservation_text.CheckState = System.Windows.Forms.CheckState.Checked;
            this.m_checkbox_with_reservation_text.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_checkbox_with_reservation_text.Location = new System.Drawing.Point(510, 165);
            this.m_checkbox_with_reservation_text.Name = "m_checkbox_with_reservation_text";
            this.m_checkbox_with_reservation_text.Size = new System.Drawing.Size(132, 20);
            this.m_checkbox_with_reservation_text.TabIndex = 19;
            this.m_checkbox_with_reservation_text.Text = "Reservationstext";
            this.m_checkbox_with_reservation_text.UseVisualStyleBackColor = true;
            // 
            // NewsletterForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(941, 718);
            this.ControlBox = false;
            this.Controls.Add(this.m_checkbox_with_reservation_text);
            this.Controls.Add(this.m_label_version);
            this.Controls.Add(this.m_checkbox_only_down_loaded_pics);
            this.Controls.Add(this.m_button_update);
            this.Controls.Add(this.m_button_edit_config);
            this.Controls.Add(this.m_button_help);
            this.Controls.Add(this.m_label_attachment);
            this.Controls.Add(this.m_combo_box_attachment);
            this.Controls.Add(this.m_picture_box);
            this.Controls.Add(this.m_combo_box_bcc_test);
            this.Controls.Add(this.m_combo_box_pics);
            this.Controls.Add(this.m_button_subject_combine);
            this.Controls.Add(this.m_combo_box_from);
            this.Controls.Add(this.m_label_band);
            this.Controls.Add(this.m_text_box_band);
            this.Controls.Add(this.m_label_zeit);
            this.Controls.Add(this.m_label_datum);
            this.Controls.Add(this.m_combo_box_minute);
            this.Controls.Add(this.m_combo_box_hour);
            this.Controls.Add(this.m_date_time_picker);
            this.Controls.Add(this.labelText);
            this.Controls.Add(this.m_button_test_send);
            this.Controls.Add(this.m_label_test_address);
            this.Controls.Add(this.m_textbox_message);
            this.Controls.Add(this.m_button_ende);
            this.Controls.Add(this.m_picture_logo);
            this.Controls.Add(this.m_label_poster);
            this.Controls.Add(this.m_label_from);
            this.Controls.Add(this.m_rich_textbox_body);
            this.Controls.Add(this.m_textbox_subject);
            this.Controls.Add(this.m_label_subject);
            this.Controls.Add(this.m_label_to);
            this.Controls.Add(this.m_button_all_send);
            this.Controls.Add(this.m_textbox_to);
            this.Controls.Add(this.m_label_colone);
            this.Controls.Add(this.m_rich_text_box_black);
            this.Font = new System.Drawing.Font("Arial", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Red;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Menu = this.mMainMenu;
            this.Name = "NewsletterForm";
            this.Text = "Newsletter";
            this.Load += new System.EventHandler(this.MainWnd_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.NewsletterForm_Paint);
            ((System.ComponentModel.ISupportInitialize)(this.m_picture_logo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_picture_box)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

    }
} // namespace
