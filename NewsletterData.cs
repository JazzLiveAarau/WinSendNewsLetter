using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace SendNewsLetter
{
    /// <summary>Holds data and constructs text blocks for the Email that shall be sent.
    /// <para>The normal use of this class is:</para>
    /// <para>* The user composes the newsletter with a Dialog. Example NewsletterForm.</para>
    /// <para>* The Dialog has a set NewsletterData function. Example NewsletterForm._GetInputFromDialog</para>
    /// <para>* The set NewsletterData function is called by a Dialog event function. Example NewsletterForm._ButtonSendAll_Click</para>
    /// <para>* The Dialog event function then calls an execution function with NewsletterData as input parameter. Example NewsletterMain.SendNewsLetterAll</para>
    /// <para>* The execution function sets email data like Subject, To, etc with get functions from NewsletterData</para>
    /// <para></para>
    /// <para></para>
    /// </summary>
    public class NewsletterData
    {
        #region Member parameters

        /// <summary>Server name</summary>
        public string m_server_name = "asmtp.mail.hostpoint.ch";

        /// <summary>Credential Email addresse</summary>
        public string m_credential_email = "jazzliv6@jazzliveaarau.ch";

        /// <summary>Credential password</summary>
        public string m_credential_passw = "SkickaAlla";

         /// <summary>FTP host name</summary>
        public string m_ftp_host = "jazzliveaarau.ch";

        /// <summary>FTP user</summary>
        public string m_ftp_user = "adressen@jazzliveaarau.ch";

        /// <summary>FTP user</summary>
        public string m_ftp_password = "aarau";

        /// <summary>From Email addresse</summary>
        public string m_from = "";

        /// <summary>To Email addresse</summary>
        public string m_to = "";

        /// <summary>Email subject</summary>
        public string m_subject = "";

        /// <summary>Email message</summary>
        public string m_message = "";

        /// <summary>Name of the poster file (picture). The file name may be empty.</summary>
        private string m_poster_file = "";

        /// <summary>Full name of the poster file (picture) with path.</summary>
        private string m_poster_path_file = "";

        /// <summary>Name of the attachment file. The file name may be empty.</summary>
        private string m_attachment_file = "";

        /// <summary>Full name of the attachment file with path.</summary>
        private string m_attachment_path_file = "";

        /// <summary>Flag telling if the mail only shall be sent to m_bcc_test as test</summary>
        public bool m_test = false;

        /// <summary>To BCC Email test addresse</summary>
        public string m_bcc_test = "";

        /// <summary>Date for the concert</summary>
        public string m_date = "";

        /// <summary>Band name</summary>
        public string m_band = "";

        /// <summary>Concert premises</summary>
        public string m_premises = "";

        /// <summary>Flag telling if the reservation text shall be part of the email</summary>
        private bool m_add_reservation_text = true;
        /// <summary>Get and set flag telling if the reservation text shall be part of the email</summary>
        public bool AddReservationText { get { return m_add_reservation_text; } set { m_add_reservation_text = value; } }

        #endregion // Member parameters

        /// <summary>Set name of the poster file (picture) with the full path. The file name may be empty.</summary>
        public bool SetPosterName(string i_poster_file, out string o_error)
        {
            o_error = "";

            if (i_poster_file.Trim() == "")
            {
                m_poster_file = "";
                m_poster_path_file = "";
                return true;
            }

            m_poster_file = i_poster_file.Trim();

            m_poster_path_file = Path.Combine(FileUtil.PosterDirectory(), m_poster_file);

            if (!File.Exists(m_poster_path_file))
            {
                m_poster_file = "";
                m_poster_path_file = "";
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoPoster;
                return false;
            }

            return true;
        }

        /// <summary>Get name of the poster file (picture) with the full path. The file name may be empty.</summary>
        public string GetPosterName()
        {
            return m_poster_path_file;
        }

        /// <summary>Set name of the attachment file with the full path. The file name may be empty.</summary>
        public bool SetAttachmentName(string i_attachment_file, out string o_error)
        {
            o_error = "";

            if (i_attachment_file.Trim() == "")
            {
                m_attachment_file = "";
                m_attachment_path_file = "";
                return true;
            }

            m_attachment_file = i_attachment_file.Trim();

            m_attachment_path_file = Path.Combine(FileUtil.AttachmentDirectory(), m_attachment_file);

            if (!File.Exists(m_attachment_path_file))
            {
                m_attachment_file = "";
                m_attachment_path_file = "";
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoAttachment;
                return false;
            }

            return true;
        }

        /// <summary>Get name of the attachment file with the full path. The file name may be empty.</summary>
        public string GetAttachmentName()
        {
            return m_attachment_path_file;
        }

        /// <summary>Change text newline to HTML newline</summary>
        private string _NewLineTextToHtml(string i_string)
        {
            string r_html_text_new_line = "";
            if (i_string.Length == 0)
                return r_html_text_new_line;

            for (int i_char = 0; i_char < i_string.Length; i_char++)
            {
                string one_char = i_string.Substring(i_char, 1);
                if (one_char == "\n")
                    r_html_text_new_line = r_html_text_new_line + "<br>";
                else
                    r_html_text_new_line = r_html_text_new_line + one_char;
            }

            return r_html_text_new_line;
        }

        /// <summary>Returns header for a HTML file (Email)</summary>
        public string Hdr()
        {
            return "<head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=ISO-8859-1\" ></head>";
        }

        /// <summary>Returns HTML paragraph start with font</summary>
        public string ParagraphStart()
        {
            return "<p><b><span lang=DE style=\'font-size:10pt;font-family:\"Arial Narrow\";mso-ansi-language:DE\'>";
        }

        /// <summary>Returns subject as HTML</summary>
        public string SubjectHtml()
        {
            return this.m_subject;
        }

        /// <summary>Returns message as HTML</summary>
        public string MessageHtml()
        {
            string ret_htm = _NewLineTextToHtml(this.m_message);
            ret_htm = ret_htm + "<br><br>";

            return ret_htm;
        }

        /* 20250221 QQQQQQQQQQQQ
        /// <summary>Returns subject as plain text</summary>
        public string SubjectTxt()
        {
            return this.m_subject;
        }
        /// <summary>Returns picture (poster) as HTML</summary>
        public string PictureHtml()
        {
            return "<img src=cid:plakat>" + "<br><br>";
        }
       
        /// <summary>Returns message as plain text</summary>
        public string MessageTxt()
        {
            return this.m_message;
        }

        /// <summary>Returns info@jazzliveaarau.ch as plain text</summary>
        public string InfoTxt()
        {
            return NewsLetterSettings.Default.EmailInfo + @" jazzliveaarau.ch";
        }
        /// <summary>Returns premises as plain text</summary>
        public string PremisesTxt()
        {
            if (this.m_premises.Length > 0)
            {
                return NewsLetterSettings.Default.PremisesLabel + @": " + this.m_premises;
            }
            else
            {
                return @"";
            }

        } // PremisesTxt
        /// <summary>Returns info@jazzliveaarau.ch as plain text</summary>
        public string ReservationInternetTxt()
        {
            return NewsLetterSettings.Default.EmailInfo + @" www.jazzliveaarau.ch/Reservation/StartReservation.htm";

        } // ReservationInternetTxt

        /// <summary>Returns reservation as HTML</summary>
        public string ReservationHtml()
        {
            string ret_htm = NewsLetterSettings.Default.EmailReservations + " ";
            ret_htm = ret_htm + "<a href=" + @"mailto" + @":reservation" + @"@jazzliveaarau.ch?" + 
                @"Subject=Reservation%20JAZZ%20live%20AARAU%20Konzert%20" +
                MailUtil.SpaceToPercentage(m_date) + @"%20" + MailUtil.SpaceToPercentage(m_band) +
                @"&body=" + MailUtil.NewLines(NewsLetterSettings.Default.EmailReservationBody, true) + MailUtil.NewLineMail + MailUtil.NewLineMail +
                MailUtil.NewLines(NewsLetterSettings.Default.EmailReservationBeforeTime, true) + MailUtil.Space +
                MailUtil.SpaceToPercentage(NewsLetterSettings.Default.EmailReservationTime) + MailUtil.Space +
                MailUtil.NewLines(NewsLetterSettings.Default.EmailReservationAfterTime, true) + MailUtil.NewLineMail + MailUtil.NewLineMail;

            ret_htm = ret_htm + @">" + NewsLetterSettings.Default.EmailAddressReservation +  @"</a>";
            if (String.Compare(NewsLetterSettings.Default.EmailReservationTelephone, "KeinTelefon", true) == 0 )
            {
                ret_htm = ret_htm + "<br>";
            }
            else
            {
                ret_htm = ret_htm + " &nbsp; oder " + NewsLetterSettings.Default.EmailReservationTelephone + ". <br>";
            }

            ret_htm = ret_htm + NewsLetterSettings.Default.EmailReservationBeforeTime + @" " + 
                                NewsLetterSettings.Default.EmailReservationTime + @" " +
                                NewsLetterSettings.Default.EmailReservationAfterTime + @" " +
                                @"<br><br>";

            return ret_htm;
        }

        /// <summary>Returns info@jazzliveaarau.ch as plain text</summary>
        public string ReservationTxt()
        {
            if (String.Compare(NewsLetterSettings.Default.EmailReservationTelephone, "KeinTelefon", true) == 0)
            {
                return NewsLetterSettings.Default.EmailReservations + " " +
                       NewsLetterSettings.Default.EmailAddressReservation + ".";
            }
            else
            {
                return NewsLetterSettings.Default.EmailReservations + " " +
                NewsLetterSettings.Default.EmailPremise + " " + NewsLetterSettings.Default.EmailAddressReservation + " oder "
                + NewsLetterSettings.Default.EmailReservationTelephone + ".";
            }
        }
 /// <summary>Returns end subscription Email adresse as plain</summary>
        public string EndSubscriptionTxt()
        {
            string ret_htm = NewsLetterSettings.Default.EmailEndSubsription + " ";
            ret_htm = ret_htm + NewsLetterSettings.Default.EmailAddressInfo;

            return ret_htm;
        }

         20250221 QQQQQQQQQQQQ */

        /// <summary>Returns info@jazzliveaarau.ch as HTML</summary>
        public string InfoHtml()
        {
            string ret_htm = NewsLetterSettings.Default.EmailInfo + " ";
            ret_htm = ret_htm + "<a href=https://jazzliveaarau.ch/  target=\"_blank\">jazzliveaarau.ch</a>";
            ret_htm = ret_htm + " &nbsp;<br><br>";

            return ret_htm;
        }

        /// <summary>Returns premises as HTML</summary>
        public string PremisesHtml()
        {
            string ret_html = @"";

            if (this.m_premises.Length > 0)
            {
                ret_html = ret_html + NewsLetterSettings.Default.PremisesLabel + @": ";

                ret_html = ret_html + this.m_premises;

                ret_html = ret_html + " &nbsp;<br><br>";
            }

            return ret_html;

        } // PremisesHtml

        /// <summary>Returns www.jazzliveaarau.ch as HTML</summary>
        public string ReservationInternetHtml()
        {
            string ret_htm = NewsLetterSettings.Default.EmailReservations + " ";
            ret_htm = ret_htm + "<a href=https://jazzliveaarau.ch  target=\"_blank\">Homepage</a>";
            ret_htm = ret_htm + " &nbsp;<br><br>";

            return ret_htm;

        } // ReservationInternetHtml

        /// <summary>Returns end subscription Email adresse as HTML</summary>
        public string EndSubscriptionHtml()
        {
            string ret_htm = NewsLetterSettings.Default.EmailEndSubsription + " ";
            ret_htm = ret_htm + "<a href=\"mailto:info@jazzliveaarau.ch?Subject=Newsletter%20Abonnement%20beenden\">info@jazzliveaarau.ch</a> ";
            ret_htm = ret_htm + "<br><br>";

            return ret_htm;
        }

        /// <summary>Returns HTML paragraph end</summary>
        public string ParagraphEnd()
        {
            return "<br></b></span></p>";
        }

        /// <summary>Set the body text</summary>
        public bool SetBodyText(string i_body, out string o_error)
        {
            o_error = "";

            m_message = i_body;

            return true;
        }

        /// <summary>Check all data</summary>
        public bool Check(out string o_error)
        {
            o_error = "";

            if (this.m_to.Trim() == "")
            {
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoTo;
                return false;
            }

            if (this.m_from.Trim() == "")
            {
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoFrom;
                return false;
            }

            if (this.m_subject.Trim() == "")
            {
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoSubject;
                return false;
            }
            if (this.m_bcc_test.Trim() == "" && this.m_test)
            {
                o_error = SendNewsLetter.NewsLetterSettings.Default.ErrMsgNoSubject;
                return false;
            }
            return true;
        }
    }
}
