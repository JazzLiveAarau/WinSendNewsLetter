using JazzApp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendNewsLetter
{
    /// <summary>Utility functions for XML data from the application XML object and the current season XML object 
    /// <para>The application XML object corresponds to XML file JazzApplications.xml</para>
    /// <para>The current season XML object corresponds to the XML file JazzProgramm_20xx_20yy.xml for the current season.</para>
    /// <para></para>
    /// <para></para>
    /// </summary>
    public static class NewsLetterXml
    {
        #region Email member variables with get functions

        /// <summary>From email addresses</summary>
        private static string[] m_from_email_addresses = null;

        /// <summary>Test email addresses</summary>
        private static string[] m_test_email_addresses = null;

        /// <summary>Returns from email addresses</summary>
        private static string[] GetFromEmailAddresses() { return m_from_email_addresses; }

        /// <summary>Returns test email addresses</summary>
        private static string[] GetTestEmailAddresses() { return m_test_email_addresses; }

        #endregion // Email member variables with get functions

        #region Concert member variables with get functions

        /// <summary>All jazz concert objects for the current season</summary>
        private static JazzConcert[] m_all_concerts = null;
        /// <summary>Get All jazz concert objects for the current season</summary>
        private static JazzConcert[] GetConcerts() { return m_all_concerts; }

        /// <summary>Next jazz concert</summary>
        private static JazzConcert m_next_concert = null;
        /// <summary>Get next concert. Note that the object may be null. No error! Only because no next concert was found</summary>
        public static JazzConcert GetNextConcert() { return m_next_concert; }

        #endregion // Concert member variables with get functions

        #region Initialization functions

        /// <summary>Initialization of the XML member variables of this class
        /// <para>1. Initialize XML objects (create from XML files). Call of JazzXml.InitApplicationAndCurrentSeasonXml.</para>
        /// <para>2. Initialize Email addresses. Call of InitMemberEmailAddresses.</para>
        /// <para>3. Initialize concert objects. Call of InitMemberConcerts.</para>
        /// <para></para>
        /// </summary>
        /// <param name="i_ftp_password">FTP password</param>
        /// <param name="i_exec_dir">Path to the local exe directory</param>
        /// <param name="o_error">Error message</param>
        static public bool Init(string i_ftp_password, string i_exec_dir, out string o_error)
        {
            o_error = @"";

            // https://stackoverflow.com/questions/19506829/reading-xml-file-from-https-url/19530834
            // Does not work System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3;
            // Does not work System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            JazzXml.SetFtpConnectionData(NewsLetterSettings.Default.FtpHost, NewsLetterSettings.Default.FtpUser, i_ftp_password, i_exec_dir);

            JazzXml.InitApplicationAndCurrentSeasonXml();

            if (!InitMemberEmailAddresses(out o_error))
            {
                o_error = @"NewsLetterXml.Init InitMemberEmailAddresses failed " + o_error;
            }

            if (!InitMemberConcerts(out o_error))
            {
                o_error = @"NewsLetterXml.Init InitMemberEmailAddresses failed " + o_error;
            }

            return true;

        } // Init

        /// <summary>Initialize Email addresses for the members in the jazz club
        /// <para>1. Get active members. Call of JazzApp.GetActiveMembers</para>
        /// <para>2. Set arrays m_from_email_addresses and m_test_email_addresses</para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        private static bool InitMemberEmailAddresses(out string o_error)
        {
            o_error = "";

            MemberData[] active_members = JazzXml.GetActiveMembers();
            if (null == active_members)
            {
                o_error = "NewsletterXml.InitMemberEmailAddresses Programming error: Failure getting member data";
                return false;
            }

            ArrayList from_addresses_string_array = new ArrayList();
            ArrayList test_addresses_string_array = new ArrayList();

            for (int index_member = 0; index_member < active_members.Length; index_member++)
            {
                MemberData current_member = active_members[index_member];

                string from_email = current_member.EmailAddress;
                string test_email = current_member.PrivateEmailAddress;

                if (from_email.Trim().Length > 0)
                    from_addresses_string_array.Add(from_email.Trim());

                if (test_email.Trim().Length > 0)
                    test_addresses_string_array.Add(test_email.Trim());
            }

            m_from_email_addresses = (string[])from_addresses_string_array.ToArray(typeof(string));

            m_test_email_addresses = (string[])test_addresses_string_array.ToArray(typeof(string));

            return true;

        } // InitMemberEmailAddresses

        /// <summary>Initialize concert member objects
        /// <para>1. </para>
        /// <para>2. </para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        private static bool InitMemberConcerts(out string o_error)
        {
            o_error = "";

            int number_concerts = JazzXml.GetNumberConcertsInCurrentDocument();
            if (number_concerts <= 0)
            {
                o_error = "NewsLetterXml.InitMemberConcerts number_concerts <= 0";
                return false;
            }

            m_all_concerts = new JazzConcert[number_concerts];

            for (int concert_number=1; concert_number<=number_concerts; concert_number++)
            {
                JazzConcert current_concert = new JazzConcert();

                current_concert.BandName = JazzXml.GetBandName(concert_number);
                current_concert.Year = JazzXml.GetYearInt(concert_number);
                current_concert.Month = JazzXml.GetMonthInt(concert_number);
                current_concert.Day = JazzXml.GetDayInt(concert_number);
                string path_file_name_poster = JazzXml.GetPosterMidSize(concert_number);
                string file_name = Path.GetFileName(path_file_name_poster);
                current_concert.Poster = file_name;
                current_concert.Premises = JazzXml.GetAddress(concert_number);

                if (!current_concert.Check(out o_error))
                {
                    o_error = "NewsLetterXml.InitMemberConcerts JazzConcert.Check failed " + o_error;
                    return false;
                }

                m_all_concerts[concert_number - 1] = current_concert; // Note -1

            } // concert_number

            if (!SetNextConcert(out o_error))
            {
                o_error = "NewsLetterXml.InitMemberConcerts SetNextConcert failed " + o_error;
                return false;
            }

            return true;

        } // InitMemberConcerts

        /// <summary>Set the GetNextConcert() object
        /// <para>It is assumed that concerts are ordered with respect to date</para>
        /// <para>The object may not be set, i.e. be null. Calling functions of GetNextConcert() must check</para>
        /// <para></para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        public static bool SetNextConcert(out string o_error)
        {
            o_error = "";

            if (null == m_all_concerts || m_all_concerts.Length == 0)
            {
                o_error = "NewsLetterXml.SetNextConcert m_all_concerts is null or has no elements";
                return false;
            }

            DateTime today_date = DateTime.Now;

            for (int index_concert = 0; index_concert < m_all_concerts.Length; index_concert++)
            {
                JazzConcert current_concert = m_all_concerts[index_concert];

                DateTime date_time_concert = new DateTime(current_concert.Year, current_concert.Month, current_concert.Day);

                int number_days = (date_time_concert - today_date).Days;
                if (number_days >= 0)
                {
                    m_next_concert = current_concert;
                    return true;
                }
            }

            // Not found
            m_next_concert = null;

            return true; // Note no error. Functions must check GetNextConcert()

        } // GetNextConcertDate


        #endregion // Initialization functions

        #region Set Email controls

        /// <summary>Set the From combobox</summary>
        public static bool SetFromComboBox(ComboBox i_combo_box_from, out string o_error)
        {
            o_error = "";

            i_combo_box_from.Items.Clear();

            i_combo_box_from.Items.Add(NewsLetterSettings.Default.EmailAddressInfo);

            i_combo_box_from.Text = NewsLetterSettings.Default.EmailAddressInfo;

            string[] from_addresses = GetFromEmailAddresses();

            if (null == from_addresses || from_addresses.Length == 0)
            {
                o_error = "NewsLetterXml.SetFromComboBox from_addresses is null or has no elements";
                return false; 
            }

            foreach (string from_address in from_addresses)
            {
                i_combo_box_from.Items.Add(from_address);
            }

            return true;

        } // SetFromComboBox

        /// <summary>Set the test address combobox</summary>
        public static bool SetTestAddressComboBox(ComboBox i_combo_box_bcc_test, out string o_error)
        {
            o_error = "";

            i_combo_box_bcc_test.Items.Clear();

            string[] test_addresses = GetTestEmailAddresses();

            i_combo_box_bcc_test.Items.Add(NewsLetterSettings.Default.EmailAddressInfo);
            i_combo_box_bcc_test.Text = NewsLetterSettings.Default.EmailAddressInfo;

            if (null == test_addresses || test_addresses.Length == 0)
            {
                o_error = "NewsLetterXml.SetFromComboBox test_addresses is null or has no elements";
                return false; 
            }

            foreach (string test_address in test_addresses)
            {
                i_combo_box_bcc_test.Items.Add(test_address);
            }

            return true;

        } // SetTestAddressComboBox

        #endregion // Set Email controls

        #region Next concert

        /// <summary>Set the date time picker to the next concert
        /// <para>1. Get date fom the next concert object. Call of GetNextConcert</para>
        /// <para>If this object is null, todays date will be set</para>
        /// </summary>
        /// <param name="i_date_time_picker">Control date time picker</param>
        /// <param name="o_error">Error message</param>
        public static bool SetDateTimePickerToNextConcertDate(DateTimePicker i_date_time_picker, out string o_error)
        {
            o_error = "";

            i_date_time_picker.Format = DateTimePickerFormat.Custom;
            i_date_time_picker.CustomFormat = "d/MM";

            DateTime today_date = DateTime.Now;

            JazzConcert next_concert = GetNextConcert();

            if (next_concert != null)
            {
                DateTime date_time_next_concert = new DateTime(next_concert.Year, next_concert.Month, next_concert.Day);
                i_date_time_picker.Value = date_time_next_concert;
            }
            else
            {
                i_date_time_picker.Value = today_date;
            }

            return true;

        } // SetDateTimePickerToNextConcertDate

        /// <summary>Set the band text box to the band for the next concert
        /// <para>1. Get date fom the next concert object. Call of GetNextConcert</para>
        /// <para>If this object is null, Band Xyz will be set</para>
        /// </summary>
        /// <param name="i_text_box_band">Text box control for band name</param>
        public static void SetBandToNextConcert(TextBox i_text_box_band)
        {
            JazzConcert next_concert = GetNextConcert();

            if (next_concert != null)
            {
                i_text_box_band.Text = next_concert.BandName;
            }
            else
            {
                i_text_box_band.Text = @"Band Xyz";
            }

        } // SetTextBoxBandToNextConcert

        /// <summary>Set the premises text box for the next concert
        /// <para>1. Get date fom the next concert object. Call of GetNextConcert</para>
        /// <para>If this object is null, the text box will be empty</para>
        /// </summary>
        /// <param name="i_text_box_premises">Text box control for premises</param>
        public static void SetPremisesToNextConcert(TextBox i_text_box_premises)
        {
            JazzConcert next_concert = GetNextConcert();

            if (next_concert != null)
            {
                i_text_box_premises.Text = next_concert.Premises;
            }
            else
            {
                i_text_box_premises.Text = @"";
            }

        } // SetPremisesToNextConcert

        /// <summary>Set the poster for the next concert
        /// <para>1. Get date fom the next concert object. Call of GetNextConcert</para>
        /// <para>If this object is null, no poster will be set</para>
        /// </summary>
        /// <param name="i_combo_box_pics">Text box control for band name</param>
        public static void SetPosterToNextConcert(ComboBox i_combo_box_pics)
        {
            JazzConcert next_concert = GetNextConcert();

            if (next_concert == null)
            {
                return;
            }

            string next_concert_poster = next_concert.Poster;

            // TODO Check that it is one of the set items for the combobox

            i_combo_box_pics.Text = next_concert_poster;

        } // SetTextBoxBandToNextConcert

        #endregion // Next concert

    } // NewsLetterXml

} // namespace
