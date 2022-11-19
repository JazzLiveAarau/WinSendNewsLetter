using System;
using System.IO;
using System.Xml.Serialization;

namespace SendNewsLetter
{
    /// <summary>Holds the settings of the application.</summary>
    public sealed class NewsLetterSettings
    {
        static NewsLetterSettings defaultSettings = new NewsLetterSettings();

        // List the application settings as public fields below. Use field initializers to define the default values
        // that are used when a setting cannot be read due to absence or settings file corruption.

        /// <summary>Name of the Excel file with Email adresses </summary>
        public string FileExcel = @"JazzLiveAdressen.csv";

        /// <summary>Name of the directory with file FileExcel</summary>
        public string DirExcel = @"Excel";

        /// <summary>Name of the directory with the posters</summary>
        public string DirPoster = @"Plakat";

        /// <summary>Name of the directory with help documents</summary>
        public string DirHelp = @"Help";

        /// <summary>Name of the directory with an installer for a new version</summary>
        public string DirNewVersion = @"NeueVersion";

        /// <summary>Name of the directory with templates for season program, posters and flyers</summary>
        public string DirTemplates = @"Vorlagen";

        /// <summary>Name of the directory with the latest version info file </summary>
        public string DirVersionInfo = @"LatestVersionInfo";

        /// <summary>Name of the help file</summary>
        public string FileHelp = @"JAZZ_live_AARAU_Newsletter.rtf";

        /// <summary>Name of the directory with the posters</summary>
        public string DirAttachment = @"Anhang";

        /// <summary>Name of the directory with the log file</summary>
        public string DirLog = @"Log";

        /// <summary>Test Email addresses file in directory DirExcel</summary>
        public string FileExcelTest = @"test_email.csv";

        /// <summary>Flag telling if the poster (picture) shall be embedded in the mail</summary>
        public bool PosterEmbedded = true;

        /// <summary>Flag telling if the poster (picture) shall be attached to the mail</summary>
        public bool PosterAttached = true;

        /// <summary>Flag telling if only the downloaded posters shall be shown. Default value at start of the application</summary>
        public bool DownloadedPostersOnly = true;

        /// <summary>Subject start text</summary>
        public string SubjectStartText = "JAZZ live AARAU Konzert";

        /// <summary>FTP host</summary>
        public string FtpHost = "www.jazzliveaarau.ch";

        /// <summary>FTP user</summary>
        public string FtpUser = "jazzliv1";

        /// <summary>Configuration XML root element</summary>
        public string ConfigRootElement = "NewsLetterSettings";
        

        #region Email adresses

        /// <summary>Email address info</summary>
        public string EmailAddressInfo = "info@jazzliveaarau.ch";

        /// <summary>Email address reservation.</summary>
        public string EmailAddressReservation = "reservation@jazzliveaarau.ch";

        #endregion

        #region Error messages
        /// <summary>Error message: Poster (picture) does not exist! Filename: </summary>
        public string ErrMsgNoPoster = "Plakat (Bild) existiert nicht! Datei: ";

        /// <summary>Error message: Attachment does not exist! Filename: </summary>
        public string ErrMsgNoAttachment = "Anhang existiert nicht! Datei: ";

        /// <summary>Error message: Email addresse 'From' cannot be empty!</summary>
        public string ErrMsgNoFrom = @"Email-Adresse 'Von' darf nicht leer sein!";

        /// <summary>Error message: Email addresse 'To' cannot be empty!</summary>
        public string ErrMsgNoTo = @"Email-Adresse 'An' darf nicht leer sein!";

        /// <summary>Error message: Email 'Subject' cannot be empty!</summary>
        public string ErrMsgNoSubject = @"Email 'Betreff' darf nicht leer sein!";

        /// <summary>Error message: Email BCC test is empty!</summary>
        public string ErrMsgNoBccTest = @"Email 'Testadresse' ist leer!";

        /// <summary>Error message: There is no input Excel file </summary>
        public string ErrMsgNoExcelFile = @"Kein Excel Datei: ";

        /// <summary>Error message: No connection to Internet is available</summary>
        public string ErrMsgNoInternetConnection = @"Keine Verbindung zu Internet ist vorhanden";

        /// <summary>Error message: Failure downloading Excel file </summary>
        public string ErrMsgNoExcelFileDownload = @"Addressen sind nicht vom Server heruntergeladen";

        /// <summary>Message: Excel file is downloaded from the server</summary>
        public string MsgExcelFileDownload = @"Addressen sind vom Server heruntergeladen";

        /// <summary>Message: Excel file, posters and attachments are downloaded from the server</summary>
        public string MsgExcelPosterFilesDownload = @"Plakate und Email-Addressen sind vom Server heruntergeladen";

        /// <summary>Error message: Failure downloading posters, attachments and new version </summary>
        public string ErrMsgFilesDownload = @"Problem beim herunterladen vom Installationsprogramm neuester Version!";

        /// <summary>Message: Installer for a new version and attachments are downloaded</summary>
        public string MsgFilesDownload = @"Installationsprogramm neuester Version dieser Applikation ist heruntergeladen." + "\n" + "\n"
            + @"Bitte diese Applikation beenden und danach die neue Version installieren mit einem " + "\n"
            + @"Doppel-Klick auf SetupJazzLiveAarauNewsletter-version-m-n.exe im Ordner NeueVersion.";

        /// <summary>Message: A new version of the application is available</summary>
        public string MsgNewVersionIsAvailable = @"Es gibt eine neue Version ";

        #endregion

        // MsgNewVersionIsAvailable


        /// <summary>Batch send message</summary>
        public string BatchSendMessage = @"Emails werden geschickt an: ";

        /// <summary>Test send message</summary>
        public string TestSendMessage = @"Emails sind nicht gesendet! Testversion von Applikation Newsletter!";

        /// <summary>End send message</summary>
        public string EndSendAllMessage = @"Emails sind gesendet und die Homepage -> Archiv -> Newsletter ist aktualisiert";

        /// <summary>End send test message</summary>
        public string EndSendTestMessage = @"Test Email ist geschickt!";
        
        /// <summary>Flag telling if there shall be debug output</summary>
        public bool DebugOutput = true;

        /// <summary>Flag telling if all Newsletters not shall be sent. Value true means test version</summary>
        public bool DoNotSendAll = false;

        #region GUI strings

        /// <summary>GUI program title</summary>
        public string GuiTextProgramTitle = @"JAZZ live AARAU Newsletter";

        /// <summary>GUI program title for a test version</summary>
        public string GuiTextProgramTitleTest = @"JAZZ live AARAU Newsletter. Testversion: Emails an alle werden nicht geschickt!";

        /// <summary>GUI help dialog title</summary>
        public string GuiHelpDialogTitle = @"JAZZ live AARAU Newsletter Hilfe";

        /// <summary>GUI help dialog exit button</summary>
        public string GuiHelpDialogExit = @"Schliessen";

        /// <summary>GUI text JAZZ live AARAU</summary>
        public string GuiTextJazzLiveAarau = @"JAZZ live AARAU";

        /// <summary>GUI default band text</summary>
        public string GuiTextBand = @"Xxxx Yyyy Quartet";

        /// <summary>GUI text concert</summary>
        public string GuiTextConcert = @"Konzert";

        /// <summary>GUI Label  date</summary>
        public string GuiLabelDate = @"Datum";

        /// <summary>GUI Label  time</summary>
        public string GuiLabelTime = @"Zeit";

        /// <summary>GUI Label  band</summary>
        public string GuiLabelBand = @"Gruppe";

        /// <summary>GUI Label  band</summary>
        public string GuiLabelPoster = @"Plakat";

        /// <summary>GUI Label  text</summary>
        public string GuiLabelText = @"Text";

        /// <summary>GUI Label  to</summary>
        public string GuiLabelTo = @"An";

        /// <summary>GUI Label  from</summary>
        public string GuiLabelFrom = @"Von";

        /// <summary>GUI Label  band</summary>
        public string GuiLabelSubject = @"Betreff";

        /// <summary>GUI Label  test address</summary>
        public string GuiLabelTestAddress = @"Test-Adresse";

        /// <summary>GUI Label  attachment</summary>
        public string GuiLabelAttachment = @"Anhang";

        /// <summary>GUI Label  version</summary>
        public string GuiLabelVersion = @"Version ";

        /// <summary>GUI Button send test</summary>
        public string GuiButtonSendTest = @"Test Newsletter schicken";

        /// <summary>GUI Button send to all</summary>
        public string GuiButtonSendToAll = @"Alle Newsletters schicken";

        /// <summary>GUI Button end application</summary>
        public string GuiButtonEnd = @"Ende";

        /// <summary>GUI Button tool</summary>
        public string GuiButtonTools = @"Tools";

        /// <summary>GUI Button write subject manually</summary>
        public string GuiButtonManual = @"M";

        /// <summary>GUI Button write subject automatically</summary>
        public string GuiButtonAutomatic = @"A";

        /// <summary>GUI Button help</summary>
        public string GuiButtonHelp = @"Hilfe";

        /// <summary>GUI Button edit config file</summary>
        public string GuiButtonEditConfig = @"Edit";

        /// <summary>GUI Button download files </summary>
        public string GuiButtonUpdate = @"Download";

        /// <summary>GUI Label check box Only downloaded</summary>
        public string GuiLabelShowOnlyDownloadedPics = @"Nur aktuelle Plakate";

        /// <summary>GUI Label check box Reservation text</summary>
        public string GuiLabelAddReservationText = @"Reservationstext";

        #endregion

        /// <summary>Email text information</summary>
        public string EmailInfo = @"Weitere Infos über unser Programm:";

        /// <summary>Email text reservations</summary>
        public string EmailReservations = @"Reservationen:";

        /// <summary>Email text reservation telephone number. No telephone number if value is KeinTelefon</summary>
        public string EmailReservationTelephone = @"KeinTelefon";

        /// <summary>Email text reservation text before EmailReservationTime</summary>
        public string EmailReservationBeforeTime = @"(Reservierte Plätze müssen";

        /// <summary>Email text reservation time prior to concert</summary>
        public string EmailReservationTime = @"10 Minuten";

        /// <summary>Email text reservation text after EmailReservationTime</summary>
        public string EmailReservationAfterTime = @"vor Konzertbeginn besetzt sein, ansonsten sie freigegeben werden können)";

        /// <summary>Email text premise (restaurant)</summary>
        public string EmailPremise = @"Spaghetti Factory Salmen";

        /// <summary>Email end subscription</summary>
        public string EmailEndSubsription = @"Möchten Sie diesen Newsletter nicht mehr erhalten, schreiben Sie uns dies bitte an:";

        /// <summary>Email text body</summary>
        public string EmailReservationBody = @"Reservation" + MailUtil.NewLineWindows + MailUtil.NewLineWindows +
                                             @"--------------------------" + MailUtil.NewLineWindows +
                                             @"Name: " + MailUtil.NewLineWindows +
                                             @"Anzahl Personen: " + MailUtil.NewLineWindows +
                                             @"Gewünschte Plätze: " + MailUtil.NewLineWindows +
                                             @"Bemerkungen: " + MailUtil.NewLineWindows +
                                             @"--------------------------" + MailUtil.NewLineWindows;


        #region Tool tips

        /// <summary>GUI Tool tip application</summary>
        public string ToolTipApplication = @"Diese Applikation holt Email-Adressen von der JAZZ live AARAU Datenbank und schickt Newsletters (als BCC)";

        /// <summary>GUI Tool tip select poster</summary>
        public string ToolTipComboBoxPics = @"Plakat (oder ein anderes Bild) wählen." + "\n" 
                                          + @"Das Bild muss im Ordner Plakat sein." + "\n"
                                          + @"Newsletter ohne Plakat oder Bild ist erlaubt.";

        /// <summary>GUI Tool tip select poster</summary>
        public string ToolTipCheckBoxPics = @"Nur (gerade) heruntergeladene Plakate oder alle Bilder im Ordner Plakat";

        /// <summary>GUI Tool tip add reservation text</summary>
        public string ToolTipCheckBoxAddReservation = @"Mit oder ohne Reservationstext";

        /// <summary>GUI Tool tip set time (minute) for start of concert</summary>
        public string ToolTipComboBoxMinute = @"Konzertstart Minute wählen";

        /// <summary>GUI Tool tip set time (hour) for start of concert</summary>
        public string ToolTipComboBoxHour = @"Konzertstart Stunde wählen";

        /// <summary>GUI Tool tip set date for start of concert</summary>
        public string ToolTipDateTimePicker = @"Konzertdatum wählen";

        /// <summary>GUI Tool tip the name of the concert</summary>
        public string ToolTipTextBoxBand = @"Name für das Konzert eingeben";

        /// <summary>GUI Tool tip manual or automatic creation of subject text</summary>
        public string ToolTipButtonCombineSubject = @"Automatisch (A) wählen: Betreff wird von Datum, Zeit und Band konstruiert." + "\n" 
                                                  + @"Manuell (M): Betreff wurde vom Benutzer eingetippt oder geändert.";

        /// <summary>GUI Tool tip for subject</summary>
        public string ToolTipTextBoxSubject = @"Dieser Text wird Betreff im Email. " + "\n"
                                           +  @"Betreff soll immer mit JAZZ live AARAU anfangen." + "\n" 
                                           +  @"Betreff darf nicht leer sein!";

        /// <summary>GUI Tool tip text</summary>
        public string ToolTipTextBoxText = @"Text mit Information über das Konzert (oder über das Saisonprogramm oder ...) hinzufügen." + "\n" 
                                         + @"Direkt eintippen oder vom Text Dokument oder Word Dokument (Rich-Text-Format) kopieren." + "\n"
                                         + @"Text darf leer sein.";

        /// <summary>GUI Tool tip picture box</summary>
        public string ToolTipPictureBox = @"Zeigt das gewählte Plakat oder das gewählte Bild im Ordner Plakat.";

        /// <summary>GUI Tool tip send a test Email</summary>
        public string ToolTipButtonSendTest = @"Schicke Newsletter an dich selbst und kontrolliere den Inhalt";

        /// <summary>GUI Tool tip send all Emails</summary>
        public string ToolTipButtonSendAll = @"Newsletter an alle - in Bündel mit 60 Empfänger - schicken." + "\n"
                                           + @"Jedes Bündel muss mit OK bestätigt werden. " + "\n"
                                           + @"Mit Abbrechen werden E-Mails nicht geschickt.";

        /// <summary>GUI Tool tip text box To</summary>
        public string ToolTipTextBoxTo = @"Diese Email-Adresse sollte immer die JAZZ live AARAU info Adresse sein.";

        /// <summary>GUI Tool tip combo box From</summary>
        public string ToolTipComboBoxFrom = @"Bitte deine Email Adresse wählen oder die JAZZ live AARAU info Adresse verwenden";

        /// <summary>GUI Tool tip test Email Address</summary>
        public string ToolTipComboBoxTestAddress = @"Bitte deine eigene Adresse wählen oder eintippen.";

        /// <summary>GUI Tool tip attachment</summary>
        public string ToolTipComboBoxAttachment = @"Anhang wählen. Zum Beispiel das Saisonsprogramm." + "\n" 
                                                + @"Anhang muss im Ordner Anhang sein";

        /// <summary>GUI Tool tip end of session</summary>
        public string ToolTipButtonEnd = @"Programm beenden";

        /// <summary>GUI Tool tip edit the configuration file</summary>
        public string ToolTipButtonEditConfig = @"Mit dieser Funktion kann alle gezeigte Texte (Anweisungen, Fehlermeldungen,...) verändert werden.";

        /// <summary>GUI Tool tip help</summary>
        public string ToolTipButtonHelp = @"Instruktionen für JAZZ live AARAU Newsletter";

        /// <summary>GUI Tool tip help dialog text</summary>
        public string ToolTipHelpTextBox = @"Text vom Hilfsdokument im Hilfeordner wird gezeigt";

        /// <summary>GUI Tool tip help dialog exit</summary>
        public string ToolTipHelpExit= @"Fenster wird geschlossen";

        /// <summary>GUI Tool tip update</summary>
        public string ToolTipButtonUpdate = @"Installationsprogramm neuester Version zum Ordner NeueVersion herunterladen. " + "\n"
           + @"Danach diese Applikation beenden und die neue Version installieren.";

        #endregion

        /// <summary>Constructor</summary>
        public NewsLetterSettings() { }

        /// <summary>Gets the default settings instance.</summary>
        /// <remarks>
        /// <para>On first access, an attempt is made to load the settings from an application-specific location. If the
        /// file is not found or corrupt, then all fields of the returned instance are set to their default values.
        /// </para>
        /// </remarks>
        internal static NewsLetterSettings Default
        {
            get { return defaultSettings; }
        }

        /// <summary>Saves all settings.</summary>
        internal void Save()
        {
            // Always existing Directory.CreateDirectory(FileUtil.GetPathToExeDirectory());

            using (FileStream fileStream = new FileStream(FileUtil.ConfigFileName(), FileMode.Create))
            using (StreamWriter streamWriter = new StreamWriter(fileStream))
            {
                new XmlSerializer(typeof(NewsLetterSettings)).Serialize(streamWriter, defaultSettings);
            }
        }

        /// <summary>Reads the configuration file and sets values in defaultSettings.</summary>
        internal void ReadFromConfigFile()
        {
            using (FileStream fileStream = new FileStream(FileUtil.ConfigFileName(), FileMode.Open, FileAccess.Read, FileShare.Read))
            using (StreamReader streamReader = new StreamReader(fileStream))
            {
                defaultSettings = (NewsLetterSettings)new XmlSerializer(typeof(NewsLetterSettings)).Deserialize(streamReader);
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        static NewsLetterSettings()
        {
            try
            {
                using (FileStream fileStream = new FileStream(FileUtil.ConfigFileName(), FileMode.Open, FileAccess.Read, FileShare.Read))
                using (StreamReader streamReader = new StreamReader(fileStream))
                {
                    defaultSettings = (NewsLetterSettings)new XmlSerializer(typeof(NewsLetterSettings)).Deserialize(streamReader);
                }
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
            catch (InvalidOperationException) { } // Thrown when there is an error in the XML document
            catch (InvalidCastException) { } // Thrown occasionally in Visual Studio when opening designer
            catch (Exception e)
            {
                using (StreamWriter w = File.AppendText(Path.Combine(FileUtil.GetPathToExeDirectory(), "NewsLetter-debug-log.txt")))
                {
                    w.WriteLine();
                    w.WriteLine(">>> Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!");
                    w.WriteLine();
                    w.WriteLine(e);
                    w.WriteLine();

                    // Close the writer and underlying file.
                    w.Close();
                }
            }
        }
    }
}

